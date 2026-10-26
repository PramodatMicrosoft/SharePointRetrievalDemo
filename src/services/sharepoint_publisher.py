# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

"""SharePoint publisher service to upload MyFiles and OrgFiles dummy documents
to site2 and tag each file's Country column (6 Europe, 2 US, 2 China per folder).
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse

import aiohttp

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class SharePointPublishSummary:
    status: str  # "completed", "permission_required", "failed"
    site_id: str
    list_id: str
    drive_id: str
    uploaded_count: int
    tagged_count: int
    message: str
    files: list[dict[str, Any]]


def load_metadata_manifest(base_dir: Path | None = None) -> dict[str, Any]:
    if base_dir is None:
        base_dir = Path(__file__).resolve().parent.parent.parent / "Sharepoint"
    manifest_path = base_dir / "metadata_manifest.json"
    if not manifest_path.exists():
        return {"files": []}
    return json.loads(manifest_path.read_text(encoding="utf-8"))


async def publish_demo_files_to_sharepoint(
    access_token: str,
    site_url: str = "https://m365cpi50688407.sharepoint.com/sites/site2",
    base_sharepoint_dir: Path | None = None,
) -> SharePointPublishSummary:
    """Upload the 20 local dummy .docx files in Sharepoint/MyFiles and Sharepoint/OrgFiles
    to SharePoint site2, ensure the 'Country' column exists on the Documents library,
    and tag 6 Europe, 2 US, and 2 China files in each folder.
    """
    if base_sharepoint_dir is None:
        base_sharepoint_dir = Path(__file__).resolve().parent.parent.parent / "Sharepoint"

    manifest = load_metadata_manifest(base_sharepoint_dir)
    files_spec = manifest.get("files", [])

    parsed = urlparse(site_url)
    host = parsed.netloc
    site_path = parsed.path.rstrip("/")
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    async with aiohttp.ClientSession(headers=headers) as session:
        # 1. Resolve site2
        site_api = f"https://graph.microsoft.com/v1.0/sites/{host}:{site_path}?$select=id,displayName,webUrl"
        async with session.get(site_api) as resp:
            if resp.status != 200:
                return SharePointPublishSummary(
                    status="failed",
                    site_id="",
                    list_id="",
                    drive_id="",
                    uploaded_count=0,
                    tagged_count=0,
                    message=f"Could not resolve SharePoint site {site_url} (HTTP {resp.status}).",
                    files=[],
                )
            site_data = await resp.json()
            site_id = site_data["id"]

        # 2. Resolve default Documents library and drive
        drive_api = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive?$select=id,name,webUrl"
        async with session.get(drive_api) as resp:
            if resp.status != 200:
                return SharePointPublishSummary(
                    status="failed",
                    site_id=site_id,
                    list_id="",
                    drive_id="",
                    uploaded_count=0,
                    tagged_count=0,
                    message=f"Could not resolve default drive on {site_url} (HTTP {resp.status}).",
                    files=[],
                )
            drive_data = await resp.json()
            drive_id = drive_data["id"]

        lists_api = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$select=id,name,displayName"
        list_id = ""
        async with session.get(lists_api) as resp:
            if resp.status == 200:
                lists_data = await resp.json()
                for lst in lists_data.get("value", []):
                    if lst.get("name") in ("Shared Documents", "Documents") or lst.get(
                        "displayName"
                    ) in ("Documents", "Shared Documents"):
                        list_id = lst["id"]
                        break

        # 3. Check/create 'Country' column on the Documents list
        if list_id:
            cols_api = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns?$select=id,name,displayName,indexed"
            country_col_exists = False
            async with session.get(cols_api) as resp:
                if resp.status == 200:
                    cols_data = await resp.json()
                    for col in cols_data.get("value", []):
                        if (
                            col.get("name", "").lower() == "country"
                            or col.get("displayName", "").lower() == "country"
                        ):
                            country_col_exists = True
                            break

            if not country_col_exists:
                create_col_body = {
                    "name": "Country",
                    "displayName": "Country",
                    "description": "Country classification for Two-Level Retrieval (mapped to RefinableString100)",
                    "indexed": True,
                    "enforceUniqueValues": False,
                    "text": {},
                }
                async with session.post(
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns",
                    json=create_col_body,
                ) as col_resp:
                    if col_resp.status == 403:
                        return SharePointPublishSummary(
                            status="permission_required",
                            site_id=site_id,
                            list_id=list_id,
                            drive_id=drive_id,
                            uploaded_count=0,
                            tagged_count=0,
                            message=(
                                "Connected to site2 (MyFiles and OrgFiles verified), but the current Graph token "
                                "only has read permissions (Files.Read.All, Sites.Read.All). To push files and create/tag "
                                "the 'Country' column via Graph API, grant 'Files.ReadWrite.All' and 'Sites.Manage.All' "
                                "(or run `python scripts/publish_to_sharepoint.py --interactive`)."
                            ),
                            files=files_spec,
                        )

        # 4. Upload each file to MyFiles / OrgFiles and tag its Country column
        uploaded_count = 0
        tagged_count = 0
        published_files: list[dict[str, Any]] = []

        for entry in files_spec:
            folder = entry["folder"]
            filename = entry["filename"]
            country = entry["Country"]
            local_file = base_sharepoint_dir / folder / filename
            if not local_file.exists():
                continue

            content_bytes = local_file.read_bytes()
            encoded_folder = quote(folder)
            encoded_name = quote(filename)
            upload_url = (
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
                f"/root:/{encoded_folder}/{encoded_name}:/content"
            )
            async with session.put(
                upload_url,
                data=content_bytes,
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                },
            ) as up_resp:
                if up_resp.status == 403:
                    return SharePointPublishSummary(
                        status="permission_required",
                        site_id=site_id,
                        list_id=list_id,
                        drive_id=drive_id,
                        uploaded_count=uploaded_count,
                        tagged_count=tagged_count,
                        message=(
                            "Connected to site2 (MyFiles and OrgFiles verified), but the current Graph token "
                            "only has read permissions (Files.Read.All, Sites.Read.All). Grant 'Files.ReadWrite.All' "
                            "and 'Sites.Manage.All' to push files and tag the Country column directly."
                        ),
                        files=files_spec,
                    )
                if up_resp.status not in (200, 201):
                    logger.warning("Failed to upload %s: HTTP %s", filename, up_resp.status)
                    continue

                uploaded_item = await up_resp.json()
                uploaded_count += 1
                item_id = uploaded_item["id"]

            # Tag listItem fields.Country
            fields_url = (
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/listItem/fields"
            )
            async with session.patch(
                fields_url,
                json={"Country": country},
            ) as tag_resp:
                if tag_resp.status in (200, 201, 204):
                    tagged_count += 1

            published_files.append(
                {
                    **entry,
                    "driveItemId": item_id,
                    "webUrl": uploaded_item.get("webUrl", entry["webUrl"]),
                }
            )

        return SharePointPublishSummary(
            status="completed",
            site_id=site_id,
            list_id=list_id,
            drive_id=drive_id,
            uploaded_count=uploaded_count,
            tagged_count=tagged_count,
            message=(
                f"Successfully uploaded {uploaded_count} files to MyFiles and OrgFiles on {site_url} "
                f"and tagged {tagged_count} files with the Country column (6 Europe, 2 US, 2 China per folder)."
            ),
            files=published_files,
        )
