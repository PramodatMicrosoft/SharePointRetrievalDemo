# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from __future__ import annotations

import asyncio
import json
import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from enum import Enum
from os import environ
from pathlib import Path
from typing import Awaitable, Callable, Literal, Mapping
from urllib.parse import unquote, urlparse

import aiohttp

logger = logging.getLogger(__name__)

SUPPORTED_COUNTRIES: tuple[str, ...] = ("Europe", "US", "China")
DEFAULT_COUNTRY_MANAGED_PROPERTY = "RefinableString100"

_COUNTRY_ALIASES: dict[str, str] = {
    "europe": "Europe",
    "eu": "Europe",
    "european union": "Europe",
    "us": "US",
    "usa": "US",
    "u.s.": "US",
    "u.s.a.": "US",
    "united states": "US",
    "united states of america": "US",
    "china": "China",
    "prc": "China",
    "cn": "China",
}


class RetrievalStatus(str, Enum):
    SUCCESS = "success"
    NOT_SIGNED_IN = "not_signed_in"
    NO_RESULTS = "no_results"
    SERVICE_UNAVAILABLE = "service_unavailable"


@dataclass(frozen=True)
class RetrievalCandidateFile:
    filename: str
    title: str
    folder: str
    country: str
    managed_property: str
    web_url: str
    kql_path: str
    doc_id: str = ""


@dataclass(frozen=True)
class RetrievalItem:
    title: str
    extract: str
    web_url: str
    country: str = ""
    folder: str = ""
    filename: str = ""
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY
    doc_id: str = ""


@dataclass(frozen=True)
class RetrievalResult:
    status: RetrievalStatus
    items: list[RetrievalItem]
    country: str = ""
    level1_filter_expression: str = ""
    level1_candidates: list[RetrievalCandidateFile] = field(default_factory=list)
    level2_filter_expression: str = ""
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY


@dataclass(frozen=True)
class RetrievalOptions:
    sharepoint_site_url: str
    maximum_number_of_results: int
    data_source: Literal["sharePoint", "oneDriveBusiness"] = "sharePoint"
    country_managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY


GetAccessToken = Callable[[], Awaitable[str | None]]
PostRetrievalRequest = Callable[[str, dict], Awaitable[tuple[int, dict]]]


def normalize_country(value: str | None) -> str | None:
    if not value:
        return None
    cleaned = value.strip().lower()
    for prefix in ("select country:", "country:", "country=", "switch to ", "search "):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix) :].strip()
    if cleaned in _COUNTRY_ALIASES:
        return _COUNTRY_ALIASES[cleaned]
    return None


def get_retrieval_options(environment: Mapping[str, str] = environ) -> RetrievalOptions:
    site_url = environment.get("RETRIEVAL_SHAREPOINT_SITE_URL", "")
    parsed_url = urlparse(site_url)
    if parsed_url.scheme != "https" or not parsed_url.netloc.endswith(".sharepoint.com"):
        raise ValueError(
            "RETRIEVAL_SHAREPOINT_SITE_URL must be an absolute HTTPS SharePoint site URL."
        )

    configured_source = environment.get("RETRIEVAL_DATA_SOURCE", "sharePoint")
    data_source: Literal["sharePoint", "oneDriveBusiness"]
    if configured_source == "sharePoint":
        data_source = "sharePoint"
    elif configured_source == "oneDriveBusiness":
        data_source = "oneDriveBusiness"
    else:
        raise ValueError(
            "RETRIEVAL_DATA_SOURCE must be sharePoint or oneDriveBusiness."
        )

    try:
        maximum_number_of_results = int(
            environment.get("RETRIEVAL_MAXIMUM_NUMBER_OF_RESULTS", "3")
        )
    except ValueError as exception:
        raise ValueError(
            "RETRIEVAL_MAXIMUM_NUMBER_OF_RESULTS must be from 1 through 25."
        ) from exception
    if not 1 <= maximum_number_of_results <= 25:
        raise ValueError("RETRIEVAL_MAXIMUM_NUMBER_OF_RESULTS must be from 1 through 25.")

    country_managed_property = (
        environment.get(
            "RETRIEVAL_COUNTRY_MANAGED_PROPERTY", DEFAULT_COUNTRY_MANAGED_PROPERTY
        ).strip()
        or DEFAULT_COUNTRY_MANAGED_PROPERTY
    )

    return RetrievalOptions(
        site_url,
        maximum_number_of_results,
        data_source,
        country_managed_property,
    )


def create_filter_expression(sharepoint_site_url: str) -> str:
    return f'path:"{sharepoint_site_url.rstrip("/")}/"'


def create_level1_filter_expression(
    sharepoint_site_url: str,
    country: str,
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY,
) -> str:
    """Build Level 1 KQL filterExpression scoping to the SharePoint site and Country managed property."""
    base_filter = create_filter_expression(sharepoint_site_url)
    return f'{base_filter} AND {managed_property}:"{country}"'


def create_level2_filter_expression(
    candidate_paths: list[str],
    sharepoint_site_url: str = "",
    country: str = "",
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY,
) -> str:
    """Build Level 2 KQL filterExpression scoping semantic retrieval strictly to Level 1 approved file paths."""
    clauses = [f'Path:"{path.strip()}"' for path in candidate_paths if path and path.strip()]
    if clauses:
        return f"({' OR '.join(clauses)})"
    if sharepoint_site_url and country:
        return create_level1_filter_expression(
            sharepoint_site_url, country, managed_property
        )
    return create_filter_expression(sharepoint_site_url) if sharepoint_site_url else ""


def get_country_candidate_files(
    country: str,
    sharepoint_site_url: str | None = None,
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY,
    base_sharepoint_dir: Path | None = None,
) -> list[RetrievalCandidateFile]:
    """Return the governed candidate files for the selected country across MyFiles and OrgFiles."""
    normalized = normalize_country(country) or country.strip()
    if base_sharepoint_dir is None:
        base_sharepoint_dir = Path(__file__).resolve().parent.parent.parent / "Sharepoint"
    manifest_path = base_sharepoint_dir / "metadata_manifest.json"
    if not manifest_path.exists():
        return []

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return []

    site_root = (sharepoint_site_url or manifest.get("siteUrl", "")).rstrip("/")
    candidates: list[RetrievalCandidateFile] = []
    for entry in manifest.get("files", []):
        if entry.get("Country", "").lower() != normalized.lower():
            continue
        folder = entry.get("folder", "MyFiles")
        filename = entry.get("filename", "")
        kql_path = (
            f"{site_root}/Shared Documents/{folder}/{filename}"
            if site_root
            else entry.get("kqlPath", "")
        )
        web_url = (
            f"{site_root}/Shared%20Documents/{folder}/{filename}"
            if site_root
            else entry.get("webUrl", "")
        )
        candidates.append(
            RetrievalCandidateFile(
                filename=filename,
                title=entry.get("title", filename),
                folder=folder,
                country=entry.get("Country", normalized),
                managed_property=managed_property,
                web_url=web_url,
                kql_path=kql_path,
                doc_id=entry.get("docId", ""),
            )
        )
    return candidates


async def retrieve_sharepoint(
    question: str,
    get_access_token: GetAccessToken,
    options: RetrievalOptions | None = None,
    post_request: PostRetrievalRequest | None = None,
) -> RetrievalResult:
    if not question.strip():
        return RetrievalResult(RetrievalStatus.NO_RESULTS, [])

    try:
        access_token = await get_access_token()
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.warning("The delegated Microsoft Graph token was not available.")
        return RetrievalResult(RetrievalStatus.NOT_SIGNED_IN, [])
    if not access_token:
        return RetrievalResult(RetrievalStatus.NOT_SIGNED_IN, [])

    options = options or get_retrieval_options()
    body = {
        "queryString": question,
        "dataSource": options.data_source,
        "filterExpression": create_filter_expression(options.sharepoint_site_url),
        "resourceMetadata": ["title", "author"],
        "maximumNumberOfResults": options.maximum_number_of_results,
    }
    try:
        status, payload = await (post_request or _post_retrieval_request)(access_token, body)
        if not 200 <= status < 300:
            logger.warning("Copilot Retrieval API returned status code %s.", status)
            return RetrievalResult(RetrievalStatus.SERVICE_UNAVAILABLE, [])
        items = map_retrieval_items(payload)
        return (
            RetrievalResult(RetrievalStatus.SUCCESS, items)
            if items
            else RetrievalResult(RetrievalStatus.NO_RESULTS, [])
        )
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.error("Copilot Retrieval API request failed.")
        return RetrievalResult(RetrievalStatus.SERVICE_UNAVAILABLE, [])


async def retrieve_two_level_sharepoint(
    question: str,
    country: str,
    get_access_token: GetAccessToken,
    options: RetrievalOptions | None = None,
    post_request: PostRetrievalRequest | None = None,
    allow_local_fallback: bool | None = None,
) -> RetrievalResult:
    """Execute Two-Level Retrieval with Microsoft 365 Copilot Retrieval API (Section 8.1):
    - Level 1: Apply country metadata filter (e.g., RefinableString100:"Europe") to identify
      and confirm eligible documents in MyFiles and OrgFiles for the chosen country.
    - Level 2: Execute semantic retrieval for the user's question strictly scoped to the
      confirmed Level 1 file paths (Path:"..." OR Path:"...").
    """
    normalized_country = normalize_country(country) or country.strip()
    options = options or get_retrieval_options()
    managed_prop = options.country_managed_property

    if not question.strip() or not normalized_country:
        return RetrievalResult(
            RetrievalStatus.NO_RESULTS,
            [],
            country=normalized_country,
            managed_property=managed_prop,
        )

    try:
        access_token = await get_access_token()
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.warning("The delegated Microsoft Graph token was not available.")
        return RetrievalResult(
            RetrievalStatus.NOT_SIGNED_IN,
            [],
            country=normalized_country,
            managed_property=managed_prop,
        )
    if not access_token:
        return RetrievalResult(
            RetrievalStatus.NOT_SIGNED_IN,
            [],
            country=normalized_country,
            managed_property=managed_prop,
        )

    if allow_local_fallback is None:
        allow_local_fallback = post_request is None

    requester = post_request or _post_retrieval_request
    level1_filter = create_level1_filter_expression(
        options.sharepoint_site_url, normalized_country, managed_prop
    )
    level1_body = {
        "queryString": f"Find briefing source documents matching the selected country scope {normalized_country}",
        "dataSource": options.data_source,
        "filterExpression": level1_filter,
        "resourceMetadata": ["title", "author", managed_prop, "Country"],
        "maximumNumberOfResults": 25,
    }

    try:
        # Level 1: Discover eligible files matching Country filter
        l1_status, l1_payload = await requester(access_token, level1_body)
        if not 200 <= l1_status < 300:
            logger.warning("Level 1 Copilot Retrieval API returned status code %s.", l1_status)
            return RetrievalResult(
                RetrievalStatus.SERVICE_UNAVAILABLE,
                [],
                country=normalized_country,
                level1_filter_expression=level1_filter,
                managed_property=managed_prop,
            )

        manifest_candidates = get_country_candidate_files(
            normalized_country,
            options.sharepoint_site_url,
            managed_prop,
        )
        level1_candidates = _extract_level1_candidates(
            l1_payload,
            normalized_country,
            managed_prop,
            manifest_candidates,
        )
        if not level1_candidates and manifest_candidates:
            level1_candidates = manifest_candidates

        candidate_paths = [c.kql_path for c in level1_candidates if c.kql_path]
        level2_filter = create_level2_filter_expression(
            candidate_paths,
            options.sharepoint_site_url,
            normalized_country,
            managed_prop,
        )

        # Level 2: Semantic search scoped only to the Level 1 eligible Country files
        level2_body = {
            "queryString": question,
            "dataSource": options.data_source,
            "filterExpression": level2_filter,
            "resourceMetadata": ["title", "author", managed_prop, "Country"],
            "maximumNumberOfResults": options.maximum_number_of_results,
        }
        l2_status, l2_payload = await requester(access_token, level2_body)
        if not 200 <= l2_status < 300:
            logger.warning("Level 2 Copilot Retrieval API returned status code %s.", l2_status)
            return RetrievalResult(
                RetrievalStatus.SERVICE_UNAVAILABLE,
                [],
                country=normalized_country,
                level1_filter_expression=level1_filter,
                level1_candidates=level1_candidates,
                level2_filter_expression=level2_filter,
                managed_property=managed_prop,
            )

        items = map_retrieval_items(
            l2_payload,
            default_country=normalized_country,
            managed_property=managed_prop,
            candidates=level1_candidates,
        )
        if not items and allow_local_fallback and level1_candidates:
            items = _retrieve_from_local_country_files(
                question=question,
                country=normalized_country,
                candidates=level1_candidates,
                maximum_results=options.maximum_number_of_results,
            )

        return RetrievalResult(
            RetrievalStatus.SUCCESS if items else RetrievalStatus.NO_RESULTS,
            items,
            country=normalized_country,
            level1_filter_expression=level1_filter,
            level1_candidates=level1_candidates,
            level2_filter_expression=level2_filter,
            managed_property=managed_prop,
        )
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.error("Two-Level Copilot Retrieval API request failed.")
        return RetrievalResult(
            RetrievalStatus.SERVICE_UNAVAILABLE,
            [],
            country=normalized_country,
            level1_filter_expression=level1_filter,
            managed_property=managed_prop,
        )


async def _post_retrieval_request(access_token: str, body: dict) -> tuple[int, dict]:
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/copilot/retrieval",
            headers={"Authorization": f"Bearer {access_token}"},
            json=body,
        ) as response:
            if not 200 <= response.status < 300:
                return response.status, {}
            return response.status, await response.json()


def _infer_folder_and_filename(web_url: str) -> tuple[str, str]:
    decoded = unquote(web_url)
    filename = decoded.rstrip("/").split("/")[-1]
    if "file=" in decoded:
        match = re.search(r"file=([^&]+)", decoded)
        if match:
            filename = unquote(match.group(1))
    folder = ""
    if "/MyFiles/" in decoded or filename.contains_myfiles if hasattr(filename, "contains_myfiles") else "_MyFiles_" in filename:
        folder = "MyFiles"
    elif "/OrgFiles/" in decoded or "_OrgFiles_" in filename:
        folder = "OrgFiles"
    return folder, filename


def _extract_level1_candidates(
    payload: dict,
    country: str,
    managed_property: str,
    manifest_candidates: list[RetrievalCandidateFile],
) -> list[RetrievalCandidateFile]:
    by_filename = {c.filename.lower(): c for c in manifest_candidates}
    discovered: list[RetrievalCandidateFile] = []
    seen_urls: set[str] = set()

    for hit in payload.get("retrievalHits", []):
        web_url = (hit.get("webUrl") or "").strip()
        if not web_url or web_url in seen_urls:
            continue
        meta = hit.get("resourceMetadata") or {}
        hit_country = (
            meta.get(managed_property)
            or meta.get("Country")
            or meta.get("country")
            or ""
        )
        folder, filename = _infer_folder_and_filename(web_url)
        manifest_match = by_filename.get(filename.lower())
        if hit_country and hit_country.lower() != country.lower():
            continue
        if manifest_match:
            seen_urls.add(web_url)
            discovered.append(manifest_match)
        else:
            title = (meta.get("title") or filename or "Briefing document").strip()
            seen_urls.add(web_url)
            discovered.append(
                RetrievalCandidateFile(
                    filename=filename or title,
                    title=title,
                    folder=folder or "Shared Documents",
                    country=country,
                    managed_property=managed_property,
                    web_url=web_url,
                    kql_path=unquote(web_url),
                )
            )
    return discovered


def map_retrieval_items(
    payload: dict,
    default_country: str = "",
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY,
    candidates: list[RetrievalCandidateFile] | None = None,
) -> list[RetrievalItem]:
    by_filename = {c.filename.lower(): c for c in (candidates or [])}
    items: list[RetrievalItem] = []
    for hit in payload.get("retrievalHits", []):
        web_url = hit.get("webUrl")
        extract = next(
            (
                item.get("text", "").strip()
                for item in hit.get("extracts", [])
                if item.get("text", "").strip()
            ),
            None,
        )
        if not web_url or not extract:
            continue
        meta = hit.get("resourceMetadata", {}) or {}
        title = meta.get("title", "").strip()
        folder, filename = _infer_folder_and_filename(web_url)
        matched_candidate = by_filename.get(filename.lower())
        country_val = (
            meta.get(managed_property)
            or meta.get("Country")
            or (matched_candidate.country if matched_candidate else default_country)
            or ""
        )
        if default_country and country_val and country_val.lower() != default_country.lower():
            continue
        items.append(
            RetrievalItem(
                title=title or (matched_candidate.title if matched_candidate else "Build session information"),
                extract=extract,
                web_url=web_url,
                country=country_val,
                folder=matched_candidate.folder if matched_candidate else folder,
                filename=matched_candidate.filename if matched_candidate else filename,
                managed_property=managed_property,
                doc_id=matched_candidate.doc_id if matched_candidate else "",
            )
        )
    return items


_STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "how",
    "in", "is", "it", "of", "on", "or", "tell", "me", "the", "to", "what",
    "when", "where", "which", "who", "why", "with", "about", "show", "give",
    "summarize", "summary", "details", "information", "document", "documents",
    "files", "file", "search", "find", "please", "can", "you", "country",
}


def _tokenize_query(text: str) -> list[str]:
    tokens = re.findall(r"[a-z0-9]{2,}", text.lower())
    filtered = [t for t in tokens if t not in _STOPWORDS]
    return filtered or tokens


def _read_docx_paragraphs(docx_path: Path) -> list[str]:
    if not docx_path.exists():
        return []
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            xml_bytes = zf.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        paragraphs: list[str] = []
        for p in root.iter(f"{ns}p"):
            texts = [t.text for t in p.iter(f"{ns}t") if t.text]
            joined = "".join(texts).strip()
            if len(joined) > 60 and not joined.startswith("DOCUMENT CLASSIFICATION METADATA"):
                paragraphs.append(joined)
        return paragraphs
    except Exception:
        return []


def _retrieve_from_local_country_files(
    question: str,
    country: str,
    candidates: list[RetrievalCandidateFile],
    maximum_results: int = 3,
    base_sharepoint_dir: Path | None = None,
) -> list[RetrievalItem]:
    """Perform Level 2 semantic matching strictly across the Level 1 confirmed Country files
    in Sharepoint/MyFiles and Sharepoint/OrgFiles when SharePoint search crawl is pending.
    """
    if base_sharepoint_dir is None:
        base_sharepoint_dir = Path(__file__).resolve().parent.parent.parent / "Sharepoint"

    query_tokens = _tokenize_query(question)
    scored: list[tuple[float, int, RetrievalCandidateFile, str]] = []

    for idx, candidate in enumerate(candidates):
        if candidate.country.lower() != country.lower():
            continue
        docx_path = base_sharepoint_dir / candidate.folder / candidate.filename
        paragraphs = _read_docx_paragraphs(docx_path)
        if not paragraphs:
            continue

        best_para = paragraphs[0]
        best_para_score = -1.0
        title_lower = f"{candidate.title} {candidate.filename}".lower()
        title_bonus = sum(2.5 for tok in query_tokens if tok in title_lower)

        for para in paragraphs:
            para_lower = para.lower()
            hits = sum(2.0 for tok in query_tokens if tok in para_lower)
            # Partial stem match bonus
            stem_hits = sum(
                0.8
                for tok in query_tokens
                if len(tok) >= 4 and tok[:4] in para_lower
            )
            score = hits + stem_hits
            if score > best_para_score:
                best_para_score = score
                best_para = para

        total_score = title_bonus + max(best_para_score, 0.5)
        scored.append((total_score, -idx, candidate, best_para))

    scored.sort(key=lambda item: (item[0], item[1]), reverse=True)

    # Ensure representation from both MyFiles and OrgFiles when multiple results requested
    selected: list[tuple[float, int, RetrievalCandidateFile, str]] = []
    if maximum_results >= 2 and len(scored) >= 2:
        myfiles_picks = [s for s in scored if s[2].folder == "MyFiles"]
        orgfiles_picks = [s for s in scored if s[2].folder == "OrgFiles"]
        if myfiles_picks and orgfiles_picks:
            selected.append(myfiles_picks[0])
            selected.append(orgfiles_picks[0])
            for s in scored:
                if len(selected) >= maximum_results:
                    break
                if s not in selected:
                    selected.append(s)
        else:
            selected = scored[:maximum_results]
    else:
        selected = scored[:maximum_results]

    return [
        RetrievalItem(
            title=cand.title,
            extract=extract,
            web_url=cand.web_url,
            country=cand.country,
            folder=cand.folder,
            filename=cand.filename,
            managed_property=cand.managed_property,
            doc_id=cand.doc_id,
        )
        for _, _, cand, extract in selected
    ]
