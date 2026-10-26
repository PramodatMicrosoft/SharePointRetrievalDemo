# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from __future__ import annotations

from .retrieval_client import (
    DEFAULT_COUNTRY_MANAGED_PROPERTY,
    RetrievalCandidateFile,
    RetrievalItem,
    RetrievalResult,
)

MYFILES_SHARE_URL = "https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgDoY09VZbkjSbhtdzGiIxDLAWqgVZdBXeYMVkMw3ugS0cM?e=QCWbRD"
ORGFILES_SHARE_URL = "https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgBiVQP4gpXzToCBCpbdqgtTAVA8za08GM8R6wQbeMDa9EE?e=4fkr7v"


def create_country_selection_card(
    pending_question: str = "",
    active_country: str = "",
) -> dict:
    body: list[dict] = [
        {
            "type": "TextBlock",
            "text": "Two-Level Retrieval — Select Country Filter",
            "weight": "Bolder",
            "size": "Medium",
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": (
                "Before retrieving SharePoint content, select a **Country** scope. "
                "The agent executes **Two-Level Retrieval** using the Microsoft 365 Copilot Retrieval API:\n"
                "1. **Level 1 (Country Filter)**: Filters `MyFiles` and `OrgFiles` using `RefinableString100:\"<Country>\"` (Column `Country`).\n"
                "2. **Level 2 (Semantic Search)**: Runs your question strictly against the confirmed Country files (`Path:\"...\" OR ...`)."
            ),
            "wrap": True,
            "spacing": "Small",
        },
        {
            "type": "FactSet",
            "facts": [
                {"title": "Europe Files", "value": "6 in MyFiles + 6 in OrgFiles (12 total)"},
                {"title": "US Files", "value": "2 in MyFiles + 2 in OrgFiles (4 total)"},
                {"title": "China Files", "value": "2 in MyFiles + 2 in OrgFiles (4 total)"},
                {"title": "Managed Property", "value": "RefinableString100 (Column: Country)"},
            ],
        },
    ]
    if active_country:
        body.append(
            {
                "type": "TextBlock",
                "text": f"Current active Country filter: **{active_country}**",
                "wrap": True,
                "spacing": "Small",
            }
        )
    if pending_question:
        body.append(
            {
                "type": "TextBlock",
                "text": f"Pending question queued: *\"{pending_question}\"* — select a Country below to run retrieval.",
                "wrap": True,
                "spacing": "Small",
            }
        )

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": body,
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Europe (12 files)",
                "data": {
                    "msteams": {"type": "imBack", "value": "Europe"},
                    "action": "select_country",
                    "country": "Europe",
                    "pending_question": pending_question,
                },
            },
            {
                "type": "Action.Submit",
                "title": "US (4 files)",
                "data": {
                    "msteams": {"type": "imBack", "value": "US"},
                    "action": "select_country",
                    "country": "US",
                    "pending_question": pending_question,
                },
            },
            {
                "type": "Action.Submit",
                "title": "China (4 files)",
                "data": {
                    "msteams": {"type": "imBack", "value": "China"},
                    "action": "select_country",
                    "country": "China",
                    "pending_question": pending_question,
                },
            },
            {
                "type": "Action.OpenUrl",
                "title": "Open MyFiles (Site2)",
                "url": MYFILES_SHARE_URL,
            },
            {
                "type": "Action.OpenUrl",
                "title": "Open OrgFiles (Site2)",
                "url": ORGFILES_SHARE_URL,
            },
        ],
    }


def create_level1_scope_card(
    country: str,
    candidates: list[RetrievalCandidateFile],
    level1_filter: str,
    managed_property: str = DEFAULT_COUNTRY_MANAGED_PROPERTY,
) -> dict:
    myfiles = [c for c in candidates if c.folder == "MyFiles"]
    orgfiles = [c for c in candidates if c.folder == "OrgFiles"]
    myfiles_lines = "\n".join(
        f"- `{c.filename}` [Country={c.country}]" for c in myfiles
    ) or "- None"
    orgfiles_lines = "\n".join(
        f"- `{c.filename}` [Country={c.country}]" for c in orgfiles
    ) or "- None"

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "text": f"Level 1 Filter Active — Country: {country}",
                "weight": "Bolder",
                "size": "Medium",
                "wrap": True,
            },
            {
                "type": "FactSet",
                "facts": [
                    {"title": "Country Column", "value": country},
                    {"title": "Managed Property", "value": f'{managed_property}:"{country}"'},
                    {"title": "Level 1 Filter", "value": level1_filter},
                    {
                        "title": "Eligible Files",
                        "value": f"{len(candidates)} files ({len(myfiles)} in MyFiles, {len(orgfiles)} in OrgFiles)",
                    },
                ],
            },
            {
                "type": "TextBlock",
                "text": f"**MyFiles ({len(myfiles)} {country} files):**\n{myfiles_lines}",
                "wrap": True,
                "spacing": "Small",
            },
            {
                "type": "TextBlock",
                "text": f"**OrgFiles ({len(orgfiles)} {country} files):**\n{orgfiles_lines}",
                "wrap": True,
                "spacing": "Small",
            },
            {
                "type": "TextBlock",
                "text": "Now type your question in chat (or pick a sample question below) to execute **Level 2 Semantic Retrieval** strictly over these confirmed Country files:",
                "wrap": True,
                "spacing": "Medium",
            },
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Ask: Trade & Tariffs",
                "data": {
                    "msteams": {
                        "type": "imBack",
                        "value": "What are the key trade agreement priorities, tariff policies, and customs rules?",
                    },
                    "action": "ask_question",
                    "country": country,
                    "question": "What are the key trade agreement priorities, tariff policies, and customs rules?",
                },
            },
            {
                "type": "Action.Submit",
                "title": "Ask: AI & Digital Rules",
                "data": {
                    "msteams": {
                        "type": "imBack",
                        "value": "Summarize the AI governance regulations, cybersecurity standards, and digital sovereignty rules.",
                    },
                    "action": "ask_question",
                    "country": country,
                    "question": "Summarize the AI governance regulations, cybersecurity standards, and digital sovereignty rules.",
                },
            },
            {
                "type": "Action.Submit",
                "title": "Ask: Clean Energy & Minerals",
                "data": {
                    "msteams": {
                        "type": "imBack",
                        "value": "What are the clean energy transition commitments and critical minerals supply chain strategies?",
                    },
                    "action": "ask_question",
                    "country": country,
                    "question": "What are the clean energy transition commitments and critical minerals supply chain strategies?",
                },
            },
            {
                "type": "Action.Submit",
                "title": "Change Country",
                "data": {
                    "msteams": {"type": "imBack", "value": "change country"},
                    "action": "change_country",
                },
            },
        ],
    }


def create_source_card(
    items: list[RetrievalItem],
    result: RetrievalResult | None = None,
) -> dict:
    body: list[dict] = []

    if result and result.country:
        myfiles_cands = [c for c in result.level1_candidates if c.folder == "MyFiles"]
        orgfiles_cands = [c for c in result.level1_candidates if c.folder == "OrgFiles"]
        cand_summary = ", ".join(
            f"{c.filename} ({c.folder}, Country={c.country})"
            for c in result.level1_candidates
        )
        body.append(
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": f"Two-Level Retrieval Verification — Country: {result.country}",
                        "weight": "Bolder",
                        "wrap": True,
                    },
                    {
                        "type": "FactSet",
                        "facts": [
                            {"title": "Selected Country", "value": result.country},
                            {
                                "title": "Level 1 Filter (4.1)",
                                "value": result.level1_filter_expression,
                            },
                            {
                                "title": "Level 1 Eligible Files",
                                "value": (
                                    f"{len(result.level1_candidates)} {result.country}-tagged files "
                                    f"({len(myfiles_cands)} MyFiles, {len(orgfiles_cands)} OrgFiles)"
                                ),
                            },
                            {
                                "title": "Confirmed Country Files",
                                "value": cand_summary or f"All {result.country} files",
                            },
                            {
                                "title": "Level 2 Scope (4.2)",
                                "value": f"Semantic search restricted strictly to {len(result.level1_candidates)} {result.country} file paths",
                            },
                        ],
                    },
                ],
            }
        )

    for item in items:
        item_elements: list[dict] = [
            {"type": "TextBlock", "text": item.title, "weight": "Bolder", "wrap": True},
        ]
        if item.country or item.folder or item.filename:
            badge_parts = []
            if item.country:
                badge_parts.append(f"Country: {item.country} ({item.managed_property}=\"{item.country}\")")
            if item.folder:
                badge_parts.append(f"Folder: {item.folder}")
            if item.filename:
                badge_parts.append(f"File: {item.filename}")
            item_elements.append(
                {
                    "type": "TextBlock",
                    "text": " | ".join(badge_parts),
                    "isSubtle": True,
                    "spacing": "None",
                    "wrap": True,
                }
            )
        item_elements.append(
            {"type": "TextBlock", "text": item.extract, "wrap": True, "spacing": "Small"}
        )
        body.append(
            {
                "type": "Container",
                "items": item_elements,
                "selectAction": {
                    "type": "Action.OpenUrl",
                    "title": "Open source",
                    "url": item.web_url,
                },
            }
        )

    card: dict = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": body,
    }
    if result and result.country:
        card["actions"] = [
            {
                "type": "Action.Submit",
                "title": "Switch Country",
                "data": {
                    "msteams": {"type": "imBack", "value": "change country"},
                    "action": "change_country",
                },
            },
            {
                "type": "Action.OpenUrl",
                "title": "Open MyFiles (Site2)",
                "url": MYFILES_SHARE_URL,
            },
            {
                "type": "Action.OpenUrl",
                "title": "Open OrgFiles (Site2)",
                "url": ORGFILES_SHARE_URL,
            },
        ]
    return card
