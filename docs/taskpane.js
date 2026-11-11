/*
 * LNU Arkiv – Outlook-tillegg som lagrer valgt e-post i SharePoint med metadata.
 *
 * Flyt:
 *  1. Leser den åpne e-posten direkte fra Outlook (innhold som .eml og metadata).
 *     Tillegget har ingen tilgang til resten av postkassen.
 *  2. Logger inn med brukerens egen konto og får tilgang til arkivsiden i SharePoint
 *     (Graph-tillatelsen Sites.Selected, som bare gjelder siden appen er gitt tilgang til).
 *  3. Laster opp .eml-filen og fyller inn metadata (avsender, mottakere, dato, retning,
 *     kommunikasjonstype, prosjekt, kommentar).
 *  4. Setter Outlook-kategorien «Arkivert i SharePoint» på e-posten, hvis brukeren har den.
 *
 * Alt går direkte mellom Outlook og Microsoft 365. Ingen data sendes til andre tjenester.
 */
/* global Office, msal */
(() => {
  "use strict";

  const CFG = window.LNU_ARKIV_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["https://graph.microsoft.com/Sites.Selected"];
  const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;     // Graph: enkel opplasting opptil 4 MB
  const CHUNK_SIZE = 10 * 320 * 1024;              // Opplastingsøkt: må være multiplum av 320 KiB
  const MAX_NAME_LENGTH = 140;

  const CATEGORY = typeof CFG.category === "string" ? CFG.category.trim() : "";

  let pca = null;          // MSAL-klient
  let sp = null;           // { siteId, libListId, driveId, typeListId, projectListId }
  let current = null;      // { token, meta, owner, existing }
  let loadToken = 0;       // hindrer at et gammelt kall overskriver et nytt når e-posten byttes
  let busy = false;

  const $ = (id) => document.getElementById(id);

  // ---------- Oppstart ----------

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Outlook) {
      showStatus("Dette tillegget fungerer bare i Outlook.", "error");
      return;
    }

    $("form").addEventListener("submit", (e) => { e.preventDefault(); archive(); });
    $("reload").addEventListener("click", () => loadLookups(true));
    $("type").addEventListener("change", updateButton);
    $("project").addEventListener("change", updateButton);

    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    } catch (e) {
      // Ikke alle klienter støtter festet oppgaverute. Da lukkes ruten når man bytter e-post.
    }

    await init();
  });

  async function init() {
    const missing = ["clientId", "tenantId"].filter((k) => !CFG[k] || String(CFG[k]).includes("FYLL"));
    if (missing.length || !CFG.sharepoint || !CFG.sharepoint.hostname) {
      showStatus("Tillegget er ikke konfigurert ennå. Kjør konfigurer.py (se README).", "error");
      return;
    }

    const req = Office.context.requirements;
    if (!req.isSetSupported("Mailbox", "1.14") || !req.isSetSupported("NestedAppAuth", "1.1")) {
      showStatus(
        "Denne Outlook-versjonen er for gammel for tillegget. " +
        "Oppdater Outlook, eller bruk ny Outlook eller Outlook på nett.",
        "error"
      );
      return;
    }

    try {
      pca = await msal.createNestablePublicClientApplication({
        auth: {
          clientId: CFG.clientId,
          authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        },
        cache: { cacheLocation: "localStorage" },
      });
    } catch (e) {
      showStatus("Klarte ikke å starte innloggingen: " + describe(e), "error");
      return;
    }

    renderItemBasics();

    try {
      await resolveSharePoint();
      await loadLookups(false);
      await loadMessage();
    } catch (e) {
      showStatus(describe(e), "error");
    }
  }

  async function onItemChanged() {
    $("comment").value = "";
    clearStatus();
    hideNotice();
    renderItemBasics();
    if (!pca || !sp) return;
    try {
      await loadMessage();
    } catch (e) {
      showStatus(describe(e), "error");
    }
  }

  // ---------- Innlogging og Graph ----------

  async function getToken() {
    const request = { scopes: SCOPES };
    try {
      const result = await pca.acquireTokenSilent(request);
      return result.accessToken;
    } catch (silentError) {
      if (silentError instanceof msal.InteractionRequiredAuthError) {
        const result = await pca.acquireTokenPopup(request);
        return result.accessToken;
      }
      throw silentError;
    }
  }

  class GraphError extends Error {
    constructor(status, detail, path) {
      super(`Graph ${status}: ${detail}`);
      this.status = status;
      this.detail = detail;
      this.path = path;
    }
  }

  async function graph(path, options = {}) {
    const { method = "GET", body, headers = {} } = options;
    const url = path.startsWith("https://") ? path : GRAPH + path;

    for (let attempt = 0; attempt < 3; attempt++) {
      const token = await getToken();
      const h = { Authorization: `Bearer ${token}`, ...headers };
      let payload = body;
      if (body !== undefined && !(body instanceof Blob) && typeof body !== "string") {
        h["Content-Type"] = "application/json";
        payload = JSON.stringify(body);
      }

      const res = await fetch(url, { method, headers: h, body: payload });

      if ((res.status === 429 || res.status === 503) && attempt < 2) {
        const wait = Math.min(parseInt(res.headers.get("Retry-After") || "2", 10), 10);
        await sleep(wait * 1000);
        continue;
      }

      if (!res.ok) {
        let detail = res.statusText;
        try {
          const j = await res.json();
          detail = (j.error && j.error.message) || detail;
        } catch (e) { /* ikke JSON */ }
        throw new GraphError(res.status, detail, path);
      }

      if (res.status === 204) return null;
      return res.json();
    }
    throw new GraphError(429, "For mange forespørsler", path);
  }

  async function graphAll(path) {
    const out = [];
    let next = path;
    while (next) {
      const page = await graph(next);
      out.push(...(page.value || []));
      next = page["@odata.nextLink"] || null;
    }
    return out;
  }

  // ---------- SharePoint ----------

  async function resolveSharePoint() {
    const c = CFG.sharepoint;
    const sitePath = c.sitePath.startsWith("/") ? c.sitePath : "/" + c.sitePath;

    let site;
    try {
      site = await graph(`/sites/${c.hostname}:${sitePath}?$select=id,webUrl`);
    } catch (e) {
      if (e.status === 404) throw new Error(`Fant ikke arkivsiden ${c.hostname}${sitePath}. Sjekk innstillingene.`);
      if (e.status === 403) {
        throw new Error(
          "Ingen tilgang til arkivsiden. Enten har appen ikke fått tilgang til siden " +
          "(steg 5 i oppsettguiden), eller du er ikke medlem av siden."
        );
      }
      throw e;
    }

    const lists = await graphAll(`/sites/${site.id}/lists?$select=id,displayName,webUrl&$top=200`);
    const find = (name) => lists.find((l) => l.displayName.toLowerCase() === String(name).toLowerCase());

    const lib = find(c.library);
    const typeList = find(c.typeList);
    const projectList = find(c.projectList);
    const notFound = [
      [lib, c.library], [typeList, c.typeList], [projectList, c.projectList],
    ].filter(([x]) => !x).map(([, name]) => `«${name}»`);
    if (notFound.length) {
      throw new Error(`Fant ikke ${notFound.join(", ")} på arkivsiden. Sjekk navnene i innstillingene.`);
    }

    const drive = await graph(`/sites/${site.id}/lists/${lib.id}/drive?$select=id`);

    sp = {
      siteId: site.id,
      libListId: lib.id,
      driveId: drive.id,
      typeListId: typeList.id,
      projectListId: projectList.id,
    };

    const link = $("library-link");
    link.href = lib.webUrl;
    link.hidden = false;
  }

  async function loadLookups(manual) {
    if (!sp) return;
    if (manual) showStatus("Laster listene …");
    try {
      const [types, projects] = await Promise.all([
        fetchLookup(sp.typeListId),
        fetchLookup(sp.projectListId),
      ]);
      fillSelect($("type"), types, "Velg type …");
      fillSelect($("project"), projects, "Velg prosjekt …");
      if (manual) showStatus("Listene er oppdatert.", "ok");
    } catch (e) {
      showStatus("Klarte ikke å laste listene: " + describe(e), "error");
    }
    updateButton();
  }

  async function fetchLookup(listId) {
    const items = await graphAll(`/sites/${sp.siteId}/lists/${listId}/items?$expand=fields&$top=500`);
    return items
      .map((i) => ({ id: i.id, title: (i.fields && i.fields.Title) || "", active: !(i.fields && i.fields.Aktiv === false) }))
      .filter((i) => i.title && i.active)
      .sort((a, b) => a.title.localeCompare(b.title, "nb"));
  }

  function fillSelect(select, options, placeholder) {
    const previous = select.value;
    select.textContent = "";
    select.append(new Option(placeholder, ""));
    for (const o of options) select.append(new Option(o.title, o.id));
    if (previous && options.some((o) => o.id === previous)) select.value = previous;
    select.disabled = false;
  }

  async function findExisting(internetMessageId) {
    if (!internetMessageId) return null;
    const value = encodeURIComponent(internetMessageId.replace(/'/g, "''"));
    try {
      const res = await graph(
        `/sites/${sp.siteId}/lists/${sp.libListId}/items` +
        `?$filter=fields/MeldingsID eq '${value}'&$select=id,webUrl,createdBy,createdDateTime&$top=1`,
        { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
      );
      return (res.value && res.value[0]) || null;
    } catch (e) {
      if (e.status === 400) {
        throw new Error("Biblioteket mangler kolonnen «MeldingsID», eller den er ikke indeksert. Se oppsettguiden.");
      }
      throw e;
    }
  }

  // ---------- Valgt e-post (leses direkte fra Outlook) ----------

  function isMessage(item) {
    return item && item.itemType === Office.MailboxEnums.ItemType.Message;
  }

  function renderItemBasics() {
    const item = Office.context.mailbox.item;
    if (!isMessage(item)) {
      $("mail-subject").textContent = "Velg en e-post";
      $("mail-meta").textContent = "";
      current = null;
      updateButton();
      return;
    }
    $("mail-subject").textContent = item.subject || "(uten emne)";
    const from = item.from ? (item.from.displayName || item.from.emailAddress) : "";
    const date = item.dateTimeCreated ? formatDateTime(item.dateTimeCreated) : "";
    $("mail-meta").textContent = [from, date].filter(Boolean).join(" · ");
  }

  async function loadMessage() {
    const token = ++loadToken;
    current = null;
    updateButton();

    const item = Office.context.mailbox.item;
    if (!isMessage(item)) return;

    const meta = {
      subject: item.subject || "",
      from: item.from || null,
      to: item.to || [],
      cc: item.cc || [],
      date: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
      internetMessageId: item.internetMessageId || "",
    };
    const owner = await getMailboxOwner(item);
    const existing = await findExisting(meta.internetMessageId);

    if (token !== loadToken) return; // brukeren har byttet e-post i mellomtiden

    current = { token, meta, owner, existing };

    if (existing) {
      showNotice(
        `Denne e-posten er allerede arkivert` +
        (existing.createdBy && existing.createdBy.user ? ` av ${escapeHtml(existing.createdBy.user.displayName)}` : "") +
        (existing.createdDateTime ? ` ${formatDate(new Date(existing.createdDateTime))}` : "") +
        `. <a href="${escapeAttr(existing.webUrl)}" target="_blank" rel="noopener">Åpne i arkivet</a>`,
        "warn"
      );
      $("save").textContent = "Arkiver likevel";
    } else {
      hideNotice();
      $("save").textContent = "Arkiver";
    }
    updateButton();
  }

  // Hvem eier postkassen e-posten ligger i? Brukes bare til å avgjøre om e-posten er inn- eller utgående.
  function getMailboxOwner(item) {
    const me = Office.context.mailbox.userProfile.emailAddress;
    return new Promise((resolve) => {
      if (typeof item.getSharedPropertiesAsync !== "function") { resolve(me); return; }
      item.getSharedPropertiesAsync((result) => {
        const owner = result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.owner;
        resolve(owner || me);
      });
    });
  }

  function getEmlBlob(item) {
    return new Promise((resolve, reject) => {
      item.getAsFileAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Outlook klarte ikke å hente e-posten: " + (result.error && result.error.message)));
          return;
        }
        try {
          const bin = atob(result.value);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          resolve(new Blob([bytes], { type: "message/rfc822" }));
        } catch (e) {
          reject(new Error("Klarte ikke å lese e-postfilen fra Outlook."));
        }
      });
    });
  }

  // Returnerer "added", "already" eller "missing" (kategorien finnes ikke i brukerens kategoriliste).
  function applyCategory(item) {
    return new Promise((resolve) => {
      if (!CATEGORY || !item.categories) { resolve("off"); return; }
      item.categories.getAsync((got) => {
        const has = got.status === Office.AsyncResultStatus.Succeeded &&
          (got.value || []).some((c) => c.displayName.toLowerCase() === CATEGORY.toLowerCase());
        if (has) { resolve("already"); return; }
        item.categories.addAsync([CATEGORY], (res) => {
          resolve(res.status === Office.AsyncResultStatus.Succeeded ? "added" : "missing");
        });
      });
    });
  }

  // ---------- Arkivering ----------

  async function archive() {
    if (busy || !current || !sp) return;
    const typeId = $("type").value;
    const projectId = $("project").value;
    if (!typeId || !projectId) {
      showStatus("Velg både type og prosjekt.", "error");
      return;
    }

    busy = true;
    updateButton();
    const job = current;
    const item = Office.context.mailbox.item;
    const comment = $("comment").value.trim();

    try {
      showStatus("Henter e-posten …");
      const blob = await getEmlBlob(item);

      showStatus(`Laster opp (${formatSize(blob.size)}) …`);
      const fileName = buildFileName(job.meta);
      const folder = (CFG.sharepoint.folder || "").replace(/^\/+|\/+$/g, "");
      const relPath = (folder ? folder + "/" : "") + fileName;
      const encPath = relPath.split("/").map(encodeURIComponent).join("/");

      const driveItem = blob.size <= SIMPLE_UPLOAD_LIMIT
        ? await graph(`/drives/${sp.driveId}/root:/${encPath}:/content?@microsoft.graph.conflictBehavior=rename`, {
            method: "PUT",
            body: blob,
            headers: { "Content-Type": "message/rfc822" },
          })
        : await uploadLarge(encPath, blob);

      showStatus("Lagrer metadata …");
      let metadataOk = true;
      try {
        const listItem = await graph(`/drives/${sp.driveId}/items/${driveItem.id}/listItem?$select=id`);
        await graph(`/sites/${sp.siteId}/lists/${sp.libListId}/items/${listItem.id}/fields`, {
          method: "PATCH",
          body: buildFields(job, typeId, projectId, comment),
        });
      } catch (e) {
        metadataOk = false;
        console.error("Metadata feilet", e);
      }

      const category = await applyCategory(item);

      const link = `<a href="${escapeAttr(driveItem.webUrl)}" target="_blank" rel="noopener">${escapeHtml(driveItem.name)}</a>`;
      const categoryNote = category === "missing"
        ? `<br><br>Tips: Opprett kategorien «${escapeHtml(CATEGORY)}» i Outlook, så blir arkiverte e-poster merket automatisk.`
        : "";

      if (metadataOk) {
        showStatus(`Arkivert som ${link}.${categoryNote}`, "ok", true);
      } else {
        showStatus(
          `E-posten er lagret som ${link}, men metadata ble ikke lagret. ` +
          `Åpne filen i SharePoint og fyll inn kommunikasjonstype og prosjekt manuelt.${categoryNote}`,
          "warn",
          true
        );
      }

      $("comment").value = "";
      if (job.token === loadToken) {
        current.existing = { webUrl: driveItem.webUrl };
        $("save").textContent = "Arkiver likevel";
      }
    } catch (e) {
      showStatus("Arkiveringen feilet: " + describe(e), "error");
    } finally {
      busy = false;
      updateButton();
    }
  }

  async function uploadLarge(encPath, blob) {
    const session = await graph(`/drives/${sp.driveId}/root:/${encPath}:/createUploadSession`, {
      method: "POST",
      body: { item: { "@microsoft.graph.conflictBehavior": "rename" } },
    });

    const size = blob.size;
    for (let start = 0; start < size; start += CHUNK_SIZE) {
      const end = Math.min(start + CHUNK_SIZE, size);
      showStatus(`Laster opp … ${Math.round((end / size) * 100)} %`);
      // Opplastings-URL-en er forhåndsautorisert. Den skal ikke ha Authorization-header.
      const res = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: { "Content-Range": `bytes ${start}-${end - 1}/${size}` },
        body: blob.slice(start, end),
      });
      if (res.status === 200 || res.status === 201) return res.json();
      if (res.status !== 202) {
        throw new GraphError(res.status, "Opplastingen ble avbrutt", "uploadSession");
      }
    }
    throw new Error("Opplastingen ble ikke fullført.");
  }

  function buildFields(job, typeId, projectId, comment) {
    const m = job.meta;
    const fromAddr = m.from ? m.from.emailAddress || "" : "";
    const outgoing = fromAddr && job.owner && fromAddr.toLowerCase() === job.owner.toLowerCase();
    const recipients = [...m.to, ...m.cc].map(formatAddress).filter(Boolean);

    const fields = {
      Title: truncate(m.subject || "(uten emne)", 255),
      Avsender: truncate(formatAddress(m.from), 255),
      Mottakere: recipients.join("\n"),
      Retning: outgoing ? "Utgående" : "Innkommende",
      KommunikasjonstypeLookupId: String(typeId),
      ProsjektLookupId: String(projectId),
      MeldingsID: truncate(m.internetMessageId || "", 255),
    };
    if (m.date && !isNaN(m.date)) fields.Mottatt = m.date.toISOString();
    if (comment) fields.Kommentar = comment;
    return fields;
  }

  function buildFileName(meta) {
    const date = meta.date && !isNaN(meta.date)
      ? meta.date.toLocaleDateString("sv-SE", { timeZone: "Europe/Oslo" })
      : "udatert";
    const sender = meta.from ? meta.from.displayName || meta.from.emailAddress || "" : "";
    const subject = meta.subject || "uten emne";
    let base = `${date} ${sender} - ${subject}`;
    base = base
      .replace(/[\u0000-\u001f"*:<>?\/\\|#%]/g, " ")   // tegn SharePoint ikke tillater
      .replace(/\s+/g, " ")
      .trim()
      .replace(/^~\$/, "")
      .replace(/[.\s]+$/, "");
    if (base.length > MAX_NAME_LENGTH) base = base.slice(0, MAX_NAME_LENGTH).trim();
    return base + ".eml";
  }

  // ---------- Hjelpefunksjoner ----------

  // Office.js: { displayName, emailAddress }
  function formatAddress(r) {
    if (!r) return "";
    const name = r.displayName;
    const address = r.emailAddress;
    if (name && address && name !== address) return `${name} <${address}>`;
    return address || name || "";
  }

  function updateButton() {
    $("save").disabled = busy || !current || !sp || !$("type").value || !$("project").value;
  }

  function showStatus(html, kind, isHtml) {
    const el = $("status");
    el.className = "status" + (kind ? ` is-${kind}` : "");
    if (isHtml) el.innerHTML = html; else el.textContent = html;
  }

  function clearStatus() { showStatus(""); }

  function showNotice(html, kind) {
    const el = $("notice");
    el.className = "notice" + (kind ? ` is-${kind}` : "");
    el.innerHTML = html;
    el.hidden = false;
  }

  function hideNotice() {
    $("notice").hidden = true;
    $("save").textContent = "Arkiver";
  }

  function describe(e) {
    if (e instanceof GraphError) {
      switch (e.status) {
        case 401: return "Innloggingen ble avvist. Lukk og åpne tillegget på nytt.";
        case 403: return "Du mangler tilgang. Sjekk at du kan redigere på arkivsiden, og at appen har fått tilgang til siden.";
        case 404: return `Fant ikke det tillegget lette etter (${e.detail}).`;
        case 409:
        case 423: return "Filen er låst eller i konflikt med en annen fil. Prøv igjen.";
        case 429: return "For mange forespørsler akkurat nå. Prøv igjen om litt.";
        default: return `Feil fra Microsoft 365 (${e.status}): ${e.detail}`;
      }
    }
    if (e && e.errorCode) return `${e.errorCode}: ${e.errorMessage || e.message}`;
    return (e && e.message) || String(e);
  }

  function formatDateTime(d) {
    return new Date(d).toLocaleString("nb-NO", { dateStyle: "medium", timeStyle: "short" });
  }

  function formatDate(d) {
    return d.toLocaleDateString("nb-NO", { dateStyle: "medium" });
  }

  function formatSize(bytes) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} kB`;
    return `${(bytes / (1024 * 1024)).toFixed(1).replace(".", ",")} MB`;
  }

  function truncate(s, n) { return s.length > n ? s.slice(0, n) : s; }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function escapeAttr(s) {
    const v = String(s || "");
    return /^https:\/\//i.test(v) ? escapeHtml(v) : "#";
  }

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }
})();
