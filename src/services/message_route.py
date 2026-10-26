# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Awaitable, Callable

from .card import (
    create_country_selection_card,
    create_level1_scope_card,
    create_source_card,
)
from .retrieval_client import (
    DEFAULT_COUNTRY_MANAGED_PROPERTY,
    GetAccessToken,
    RetrievalOptions,
    RetrievalResult,
    RetrievalStatus,
    SUPPORTED_COUNTRIES,
    create_level1_filter_expression,
    get_country_candidate_files,
     get_retrieval_options,
    normalize_country,
    retrieve_two_level_sharepoint,
)
from .sharepoint_publisher import publish_demo_files_to_sharepoint

BUILD_GENIE_RESPONSES = {
    RetrievalStatus.NOT_SIGNED_IN: "Please sign in to Microsoft 365, then select a Country (Europe, US, or China) and ask your question again.",
    RetrievalStatus.NO_RESULTS: "I couldn't find matching information in the selected Country's SharePoint files. Check the site URL, Country metadata tags (RefinableString100), and indexing, then try a more specific question.",
    RetrievalStatus.SERVICE_UNAVAILABLE: "I couldn't retrieve SharePoint information right now. Please try again later.",
}


@dataclass
class ConversationRetrievalState:
    selected_country: str = ""
    pending_question: str = ""


_CONVERSATION_STATES: dict[str, ConversationRetrievalState] = {}


def get_conversation_state(conversation_id: str) -> ConversationRetrievalState:
    if conversation_id not in _CONVERSATION_STATES:
        _CONVERSATION_STATES[conversation_id] = ConversationRetrievalState()
    return _CONVERSATION_STATES[conversation_id]


def clear_conversation_states() -> None:
    _CONVERSATION_STATES.clear()


def _parse_inline_country_and_question(text: str) -> tuple[str | None, str]:
    """Check if the user specified both a country prefix and a question, e.g. 'Europe: What are the tariffs?'."""
    stripped = text.strip()
    match = re.match(
        r"^(?:\[(?P<b_country>[^\]]+)\]|(?P<p_country>Europe|EU|US|USA|United States|China|PRC|CN))\s*[:\-–]\s*(?P<q>.+)$",
        stripped,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match:
        raw_country = match.group("b_country") or match.group("p_country")
        country = normalize_country(raw_country)
        question = (match.group("q") or "").strip()
        if country and question:
            return country, question
    return None, stripped


def grounded_answer(result: RetrievalResult) -> str:
    if result.country:
        myfiles_cands = [c for c in result.level1_candidates if c.folder == "MyFiles"]
        orgfiles_cands = [c for c in result.level1_candidates if c.folder == "OrgFiles"]
        eligible_list = ", ".join(
            f"`{c.filename}` ({c.folder}, Country={c.country})"
            for c in result.level1_candidates
        )
        sources = "\n\n".join(
            (
                f"{item.title} "
                f"[Country={item.country or result.country} | {item.managed_property}=\"{item.country or result.country}\" | Folder={item.folder or 'SharePoint'} | File={item.filename}]\n"
                f"{item.extract}\n"
                f"Source: {item.web_url}"
            )
            for item in result.items
        )
        return (
            f"Here is what I found in the configured SharePoint site for **Country = {result.country}** "
            f"(using Two-Level Retrieval):\n\n"
            f"**Level 1 — Country Filter Discovery (4.1):**\n"
            f"- Filter: `{result.level1_filter_expression}`\n"
            f"- Confirmed Eligible `{result.country}` Files ({len(result.level1_candidates)} files: "
            f"{len(myfiles_cands)} in `MyFiles`, {len(orgfiles_cands)} in `OrgFiles`): {eligible_list}\n\n"
            f"**Level 2 — Semantic Search Results from Confirmed `{result.country}` Files Only (4.2):**\n\n"
            f"{sources}"
        )

    sources = "\n\n".join(
        f"{item.title}\n{item.extract}\nSource: {item.web_url}" for item in result.items
    )
    return f"Here is what I found in the configured SharePoint site:\n\n{sources}"


async def handle_build_genie_message(
    result: RetrievalResult,
    send_text: Callable[[str], Awaitable[object]],
    send_source_card: Callable[[dict], Awaitable[object]],
) -> None:
    if result.status is not RetrievalStatus.SUCCESS:
        await send_text(BUILD_GENIE_RESPONSES[result.status])
        return
    await send_text(grounded_answer(result))
    await send_source_card(create_source_card(result.items, result))


async def handle_two_level_conversation_turn(
    conversation_id: str,
    text: str | None,
    value: dict[str, Any] | None,
    get_access_token: GetAccessToken,
    send_text: Callable[[str], Awaitable[object]],
    send_card: Callable[[dict], Awaitable[object]],
    options: RetrievalOptions | None = None,
) -> RetrievalResult | None:
    """Orchestrate the conversational Web Chat flow for Two-Level Retrieval:
    1. Prompt user to select a Country (Europe, US, China) before retrieval.
    2. When Country is selected, display Level 1 Country filter & eligible files in MyFiles/OrgFiles
       (or immediately run Level 1 + Level 2 if a pending question exists).
    3. When a question is asked with an active Country, execute Level 1 (Country filter) +
       Level 2 (Semantic search restricted to those Country files) and display results.
    """
    state = get_conversation_state(conversation_id)
    options = options or get_retrieval_options()
    raw_text = (text or "").strip()
    payload = value if isinstance(value, dict) else {}

    action = str(payload.get("action", "")).strip().lower()
    if action == "change_country" or raw_text.lower() in {
        "change country",
        "switch country",
        "select country",
        "reset",
        "country",
        "menu",
        "help",
    }:
        state.selected_country = ""
        await send_text(
            "Please select a **Country** (`Europe`, `US`, or `China`) to scope the Two-Level Retrieval search across `MyFiles` and `OrgFiles`:"
        )
        await send_card(
            create_country_selection_card(
                pending_question=state.pending_question,
                active_country="",
            )
        )
        return None

    if raw_text.lower() in {"sync sharepoint", "publish sharepoint"}:
        token = await get_access_token()
        if not token:
            await send_text(BUILD_GENIE_RESPONSES[RetrievalStatus.NOT_SIGNED_IN])
            return None
        summary = await publish_demo_files_to_sharepoint(
            access_token=token,
            site_url=options.sharepoint_site_url,
        )
        await send_text(summary.message)
        return None

    # Check if card action selected a country or asked a sample question
    chosen_country: str | None = None
    question_to_run: str = ""

    if action == "select_country":
        chosen_country = normalize_country(str(payload.get("country", "")))
        if payload.get("pending_question") and not state.pending_question:
            state.pending_question = str(payload["pending_question"]).strip()
    elif action == "ask_question":
        chosen_country = normalize_country(str(payload.get("country", ""))) or state.selected_country
        question_to_run = str(payload.get("question", "")).strip()
    else:
        direct_country = normalize_country(raw_text)
        if direct_country:
            chosen_country = direct_country
        else:
            inline_country, remaining_q = _parse_inline_country_and_question(raw_text)
            if inline_country:
                chosen_country = inline_country
                question_to_run = remaining_q
            else:
                question_to_run = raw_text

    # Case 1: User selected a Country
    if chosen_country and not question_to_run:
        state.selected_country = chosen_country
        if state.pending_question:
            question_to_run = state.pending_question
            state.pending_question = ""
        else:
            candidates = get_country_candidate_files(
                chosen_country,
                options.sharepoint_site_url,
                options.country_managed_property,
            )
            level1_filter = create_level1_filter_expression(
                options.sharepoint_site_url,
                chosen_country,
                options.country_managed_property,
            )
            await send_text(
                f"Country filter set to **{chosen_country}** (`{options.country_managed_property}:\"{chosen_country}\"`). "
                f"Level 1 filter identified **{len(candidates)} eligible `{chosen_country}` files** across `MyFiles` and `OrgFiles`. "
                f"Please ask your question to run Level 2 semantic retrieval over these `{chosen_country}` files."
            )
            await send_card(
                create_level1_scope_card(
                    country=chosen_country,
                    candidates=candidates,
                    level1_filter=level1_filter,
                    managed_property=options.country_managed_property,
                )
            )
            return None

    if chosen_country:
        state.selected_country = chosen_country

    # Case 2: User asked a question before selecting a Country -> require Country selection first!
    if not state.selected_country:
        if question_to_run:
            state.pending_question = question_to_run
            await send_text(
                f"Before retrieving data for *\"{question_to_run}\"*, please select which **Country** (`Europe`, `US`, or `China`) to search from:"
            )
        else:
            await send_text(
                "Please select a **Country** (`Europe`, `US`, or `China`) to filter documents in `MyFiles` and `OrgFiles` before running semantic retrieval:"
            )
        await send_card(
            create_country_selection_card(
                pending_question=state.pending_question,
                active_country="",
            )
        )
        return None

    # Case 3: Active Country is selected and we have a question -> Execute Two-Level Retrieval!
    result = await retrieve_two_level_sharepoint(
        question=question_to_run,
        country=state.selected_country,
        get_access_token=get_access_token,
        options=options,
    )
    await handle_build_genie_message(result, send_text, send_card)
    return result
