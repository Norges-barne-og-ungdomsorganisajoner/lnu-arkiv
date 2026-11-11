#!/usr/bin/env python3
"""
Lager manifest.xml og docs/config.js fra innstillinger.json.

Bruk:
    python3 konfigurer.py

Kjør skriptet på nytt hver gang du endrer innstillinger.json.
Tilleggets ID genereres første gang og lagres i innstillinger.json. Ikke endre den
etterpå, ellers ser Microsoft 365 det som et nytt tillegg.
"""
import json
import re
import sys
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SETTINGS = ROOT / "innstillinger.json"


def fail(msg: str) -> None:
    print(f"Feil: {msg}")
    sys.exit(1)


def main() -> None:
    s = json.loads(SETTINGS.read_text(encoding="utf-8"))

    placeholders = [k for k in ("baseUrl", "clientId", "tenantId") if "FYLL-INN" in str(s.get(k, ""))]
    if placeholders:
        fail(f"Fyll inn {', '.join(placeholders)} i innstillinger.json først.")

    base = s["baseUrl"].rstrip("/")
    if not base.startswith("https://"):
        fail("baseUrl må starte med https://")

    guid = re.compile(r"^[0-9a-fA-F-]{36}$")
    for key in ("clientId", "tenantId"):
        if not guid.match(s[key]):
            fail(f"{key} ser ikke ut som en GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).")

    if not re.match(r"^\d+\.\d+\.\d+\.\d+$", s.get("version", "")):
        fail("version må ha formatet 1.0.0.0")

    if not s.get("addinId"):
        s["addinId"] = str(uuid.uuid4())
        SETTINGS.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"Ny tilleggs-ID generert og lagret: {s['addinId']}")

    manifest = (ROOT / "maler" / "manifest.template.xml").read_text(encoding="utf-8")
    manifest = (manifest.replace("{{ADDIN_ID}}", s["addinId"])
                        .replace("{{VERSION}}", s["version"])
                        .replace("{{BASE_URL}}", base))
    (ROOT / "manifest.xml").write_text(manifest, encoding="utf-8")

    config = {
        "clientId": s["clientId"],
        "tenantId": s["tenantId"],
        "sharepoint": s["sharepoint"],
        "category": s.get("category") or None,
    }
    js = (ROOT / "maler" / "config.template.js").read_text(encoding="utf-8")
    js = js.replace("{{CONFIG_JSON}}", json.dumps(config, indent=2, ensure_ascii=False))
    (ROOT / "docs" / "config.js").write_text(js, encoding="utf-8")

    origin = re.match(r"^(https://[^/]+)", base).group(1)
    print("Ferdig. Skrev manifest.xml og docs/config.js.")
    print()
    print("Omdirigerings-URI for appregistreringen i Entra (type: Single-page application):")
    print(f"    brk-multihub://{origin.removeprefix('https://')}")

    sp = s["sharepoint"]
    site_path = sp["sitePath"] if sp["sitePath"].startswith("/") else "/" + sp["sitePath"]
    grant = {
        "roles": ["write"],
        "grantedToIdentities": [
            {"application": {"id": s["clientId"], "displayName": "LNU Arkiv (Outlook)"}}
        ],
    }
    print()
    print("Gi appen tilgang til arkivsiden (steg 5 i README), i Graph Explorer:")
    print("  1) GET  https://graph.microsoft.com/v1.0/sites/"
          f"{sp['hostname']}:{site_path}?$select=id")
    print("  2) POST https://graph.microsoft.com/v1.0/sites/<id fra steg 1>/permissions")
    print("     med denne teksten i Request body:")
    print("     " + json.dumps(grant, ensure_ascii=False))


if __name__ == "__main__":
    main()
