# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

"""CLI utility to upload the 20 generated dummy .docx files in Sharepoint/MyFiles
and Sharepoint/OrgFiles to SharePoint site2, ensure the 'Country' column exists,
and tag each file with Country = Europe (6), US (2), or China (2) per folder.

Usage:
  1. Using an interactive browser login to tenant m365cpi50688407.onmicrosoft.com:
     python scripts/publish_to_sharepoint.py --interactive

  2. Using a supplied Graph access token (with Files.ReadWrite.All and Sites.Manage.All):
     python scripts/publish_to_sharepoint.py --access-token <BEARER_TOKEN>

  3. Using the active Bot Framework delegated token (verifies site2 / uploads if write scopes granted):
     python scripts/publish_to_sharepoint.py
"""

from __future__ import annotations

import argparse
import asyncio
import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.services.sharepoint_publisher import publish_demo_files_to_sharepoint  # noqa: E402

SPO_TENANT_ID = "36eaf72e-29e9-4351-9acd-79d78196dd9b"
SPO_SITE_URL = "https://m365cpi50688407.sharepoint.com/sites/site2"


def _load_env_dict(env_path: Path) -> dict[str, str]:
    env: dict[str, str] = {}
    if not env_path.exists():
        return env
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def _acquire_bot_cached_graph_token(env: dict[str, str], user_id: str) -> str | None:
    client_id = env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID")
    client_secret = env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET")
    tenant_id = env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID")
    if not (client_id and client_secret and tenant_id):
        return None

    data = urllib.parse.urlencode(
        {
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://api.botframework.com/.default",
            "grant_type": "client_credentials",
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data=data,
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        bf_token = json.loads(resp.read().decode())["access_token"]

    url = (
        f"https://api.botframework.com/api/usertoken/GetToken"
        f"?userId={urllib.parse.quote(user_id)}&connectionName=graph&channelId=webchat"
    )
    r = urllib.request.Request(url, headers={"Authorization": f"Bearer {bf_token}"})
    with urllib.request.urlopen(r, timeout=10) as uresp:
        return json.loads(uresp.read().decode()).get("token")


def _acquire_interactive_token(client_id: str, tenant_id: str) -> str:
    import msal

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    result = app.acquire_token_interactive(
        scopes=["Files.ReadWrite.All", "Sites.Manage.All", "Sites.ReadWrite.All"],
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Interactive authentication failed: {result.get('error_description') or result.get('error')}"
        )
    return result["access_token"]


async def _main() -> int:
    parser = argparse.ArgumentParser(
        description="Publish MyFiles and OrgFiles dummy .docx files and Country tags to SharePoint site2."
    )
    parser.add_argument("--site-url", default=SPO_SITE_URL, help="Target SharePoint site URL")
    parser.add_argument("--access-token", default="", help="Delegated Microsoft Graph bearer token")
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Acquire a write-capable Graph token interactively via browser",
    )
    parser.add_argument(
        "--client-id",
        default="14d82eec-204b-4c2f-b7e8-296a70dab67e",
        help="Public client ID for interactive auth (defaults to Microsoft Graph CLI)",
    )
    parser.add_argument(
        "--tenant-id",
        default=SPO_TENANT_ID,
        help="Microsoft Entra tenant ID for SharePoint site2",
    )
    parser.add_argument(
        "--user-id",
        default="843fe1fe-8060-4326-9149-df2bedc9203a",
        help="Web Chat user ID for Bot Framework cached token lookup",
    )
    args = parser.parse_args()

    token = args.access_token.strip()
    if not token and args.interactive:
        token = _acquire_interactive_token(args.client_id, args.tenant_id)
    if not token:
        env = _load_env_dict(ROOT_DIR / ".env")
        token = _acquire_bot_cached_graph_token(env, args.user_id) or ""

    if not token:
        print("No Graph access token available. Run with --interactive or --access-token.")
        return 1

    summary = await publish_demo_files_to_sharepoint(
        access_token=token,
        site_url=args.site_url,
        base_sharepoint_dir=ROOT_DIR / "Sharepoint",
    )
    print(f"Status: {summary.status}")
    print(f"Site ID: {summary.site_id}")
    print(f"List ID: {summary.list_id}")
    print(f"Drive ID: {summary.drive_id}")
    print(f"Uploaded: {summary.uploaded_count} | Tagged: {summary.tagged_count}")
    print(f"Message: {summary.message}")
    return 0 if summary.status in ("completed", "permission_required") else 1


if __name__ == "__main__":
    raise SystemExit(asyncio.run(_main()))
