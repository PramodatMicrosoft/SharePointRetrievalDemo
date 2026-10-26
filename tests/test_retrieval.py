# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET
import zipfile

from src.services.message_route import (
    BUILD_GENIE_RESPONSES,
    clear_conversation_states,
    handle_build_genie_message,
    handle_two_level_conversation_turn,
)
from src.services.retrieval_client import (
    RetrievalItem,
    RetrievalOptions,
    RetrievalResult,
    RetrievalStatus,
    create_filter_expression,
    create_level1_filter_expression,
    create_level2_filter_expression,
    get_country_candidate_files,
    get_retrieval_options,
    retrieve_sharepoint,
    retrieve_two_level_sharepoint,
)


OPTIONS = RetrievalOptions("https://contoso.sharepoint.com/sites/Build", 3)


class RetrievalTests(unittest.IsolatedAsyncioTestCase):
    def test_creates_configured_site_filter(self):
        self.assertEqual(
            create_filter_expression(OPTIONS.sharepoint_site_url),
            'path:"https://contoso.sharepoint.com/sites/Build/"',
        )
        with self.assertRaises(ValueError):
            get_retrieval_options({"RETRIEVAL_SHAREPOINT_SITE_URL": "http://contoso.sharepoint.com/sites/Build"})

    def test_validates_configured_data_source_and_preserves_default(self):
        environment = {"RETRIEVAL_SHAREPOINT_SITE_URL": OPTIONS.sharepoint_site_url}
        self.assertEqual(get_retrieval_options(environment).data_source, "sharePoint")
        environment["RETRIEVAL_DATA_SOURCE"] = "oneDriveBusiness"
        self.assertEqual(
            get_retrieval_options(environment).data_source, "oneDriveBusiness"
        )
        for source in ("", "onedrivebusiness", "externalItem"):
            with self.subTest(source=source):
                environment["RETRIEVAL_DATA_SOURCE"] = source
                with self.assertRaisesRegex(ValueError, "RETRIEVAL_DATA_SOURCE"):
                    get_retrieval_options(environment)

    async def test_retrieves_mapped_items_with_a_delegated_token(self):
        captured: dict = {}

        async def post_request(token: str, body: dict):
            captured["token"] = token
            captured["body"] = body
            return 200, {"retrievalHits": [{"webUrl": "https://contoso.sharepoint.com/session.docx", "extracts": [{"text": "Session details"}], "resourceMetadata": {"title": "Pricing Analytics"}}]}

        result = await retrieve_sharepoint("Pricing Analytics", lambda: _token("delegated-token"), OPTIONS, post_request)
        self.assertEqual(result.status, RetrievalStatus.SUCCESS)
        self.assertEqual(result.items[0].title, "Pricing Analytics")
        self.assertEqual(captured["token"], "delegated-token")
        self.assertEqual(captured["body"]["dataSource"], "sharePoint")
        self.assertEqual(
            captured["body"]["filterExpression"],
            create_filter_expression(OPTIONS.sharepoint_site_url),
        )

    async def test_configured_source_preserves_question_and_site_scope(self):
        captured: dict = {}
        options = get_retrieval_options({
            "RETRIEVAL_SHAREPOINT_SITE_URL": OPTIONS.sharepoint_site_url,
            "RETRIEVAL_DATA_SOURCE": "oneDriveBusiness",
        })

        async def post_request(_token: str, body: dict):
            captured.update(body)
            return 200, {"retrievalHits": [{
                "webUrl": OPTIONS.sharepoint_site_url + "/Shared%20Documents/session.docx",
                "extracts": [{"text": "Session details"}],
                "resourceMetadata": {"title": "Pricing Analytics"},
            }]}

        result = await retrieve_sharepoint(
            "Pricing Analytics", lambda: _token("delegated-token"), options, post_request
        )
        self.assertEqual(result.status, RetrievalStatus.SUCCESS)
        self.assertEqual(captured["dataSource"], "oneDriveBusiness")
        self.assertEqual(captured["queryString"], "Pricing Analytics")
        self.assertEqual(
            captured["filterExpression"],
            create_filter_expression(OPTIONS.sharepoint_site_url),
        )
        self.assertEqual(
            captured["maximumNumberOfResults"], OPTIONS.maximum_number_of_results
        )

    async def test_maps_token_and_service_failures_to_safe_statuses(self):
        async def no_token():
            raise ValueError("secret")

        async def failed_request(_token: str, _body: dict):
            return 502, {}

        self.assertEqual((await retrieve_sharepoint("Build", no_token, OPTIONS)).status, RetrievalStatus.NOT_SIGNED_IN)
        self.assertEqual((await retrieve_sharepoint("Build", lambda: _token("token"), OPTIONS, failed_request)).status, RetrievalStatus.SERVICE_UNAVAILABLE)
        self.assertEqual((await retrieve_sharepoint("Build", lambda: _token("token"), OPTIONS, lambda _token, _body: _response(200, {"retrievalHits": []}))).status, RetrievalStatus.NO_RESULTS)

    async def test_propagates_cancellation(self):
        async def cancelled_token():
            raise asyncio.CancelledError

        with self.assertRaises(asyncio.CancelledError):
            await retrieve_sharepoint("Build", cancelled_token, OPTIONS)

        async def cancelled_request(_token: str, _body: dict):
            raise asyncio.CancelledError

        with self.assertRaises(asyncio.CancelledError):
            await retrieve_sharepoint(
                "Build", lambda: _token("token"), OPTIONS, cancelled_request
            )

    async def test_message_route_sends_grounded_text_and_source_card(self):
        messages: list[str] = []
        cards: list[dict] = []
        result = RetrievalResult(RetrievalStatus.SUCCESS, [RetrievalItem("Pricing Analytics", "Session details", "https://contoso.sharepoint.com/session.docx")])
        await handle_build_genie_message(result, lambda text: _append(messages, text), lambda card: _append(cards, card))
        self.assertIn("https://contoso.sharepoint.com/session.docx", messages[0])
        self.assertEqual(len(cards), 1)
        self.assertIn("https://contoso.sharepoint.com/session.docx", str(cards[0]))

    async def test_message_route_sends_safe_messages_for_failures(self):
        for status in (RetrievalStatus.NOT_SIGNED_IN, RetrievalStatus.NO_RESULTS, RetrievalStatus.SERVICE_UNAVAILABLE):
            messages: list[str] = []
            cards: list[dict] = []
            await handle_build_genie_message(RetrievalResult(status, []), lambda text: _append(messages, text), lambda card: _append(cards, card))
            self.assertEqual(messages, [BUILD_GENIE_RESPONSES[status]])
            self.assertEqual(cards, [])

    def test_dummy_files_in_myfiles_and_orgfiles_have_country_tags_and_six_pages(self):
        sharepoint_dir = Path(__file__).resolve().parent.parent / "Sharepoint"
        for folder in ("MyFiles", "OrgFiles"):
            folder_path = sharepoint_dir / folder
            self.assertTrue(folder_path.is_dir(), f"Missing folder {folder_path}")
            docx_files = sorted(folder_path.glob("*.docx"))
            self.assertEqual(len(docx_files), 10)

            country_counts = {"Europe": 0, "US": 0, "China": 0}
            for docx_file in docx_files:
                with zipfile.ZipFile(docx_file, "r") as zf:
                    custom_xml = ET.fromstring(zf.read("docProps/custom.xml"))
                    doc_xml = ET.fromstring(zf.read("word/document.xml"))

                props = {
                    el.attrib.get("name"): "".join(el.itertext()).strip()
                    for el in custom_xml
                    if el.attrib.get("name")
                }
                self.assertIn("Country", props)
                self.assertEqual(props["Country"], props.get("RefinableString100"))
                country_counts[props["Country"]] += 1

                # Verify >= 5 pages (5 explicit page breaks -> 6 pages)
                ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                page_breaks = [
                    br
                    for br in doc_xml.iter(f"{ns}br")
                    if br.attrib.get(f"{ns}type") == "page"
                ]
                self.assertGreaterEqual(len(page_breaks), 5)

            self.assertEqual(country_counts, {"Europe": 6, "US": 2, "China": 2})

    def test_creates_level1_and_level2_filter_expressions(self):
        l1 = create_level1_filter_expression(
            OPTIONS.sharepoint_site_url, "Europe", "RefinableString100"
        )
        self.assertEqual(
            l1,
            'path:"https://contoso.sharepoint.com/sites/Build/" AND RefinableString100:"Europe"',
        )
        l2 = create_level2_filter_expression(
            [
                "https://contoso.sharepoint.com/sites/Build/Shared Documents/MyFiles/EU_01.docx",
                "https://contoso.sharepoint.com/sites/Build/Shared Documents/OrgFiles/EU_01.docx",
            ]
        )
        self.assertEqual(
            l2,
            '(Path:"https://contoso.sharepoint.com/sites/Build/Shared Documents/MyFiles/EU_01.docx" OR Path:"https://contoso.sharepoint.com/sites/Build/Shared Documents/OrgFiles/EU_01.docx")',
        )

    async def test_two_level_retrieval_executes_level1_filter_then_level2_semantic_search(self):
        calls: list[dict] = []

        async def post_request(_token: str, body: dict):
            calls.append(body)
            if len(calls) == 1:
                # Level 1 discovery response
                return 200, {
                    "retrievalHits": [
                        {
                            "webUrl": "https://contoso.sharepoint.com/sites/Build/Shared%20Documents/MyFiles/EU_01_MyFiles_Trade_And_CBAM_Tariff_Briefing.docx",
                            "extracts": [{"text": "CBAM scope"}],
                            "resourceMetadata": {
                                "title": "Europe Bilateral Trade & CBAM",
                                "RefinableString100": "Europe",
                            },
                        },
                        {
                            "webUrl": "https://contoso.sharepoint.com/sites/Build/Shared%20Documents/OrgFiles/EU_01_OrgFiles_European_Union_Comprehensive_Trade_Policy_Framework.docx",
                            "extracts": [{"text": "CETA scope"}],
                            "resourceMetadata": {
                                "title": "Europe Institutional Policy Framework",
                                "RefinableString100": "Europe",
                            },
                        },
                    ]
                }
            # Level 2 semantic response
            return 200, {
                "retrievalHits": [
                    {
                        "webUrl": "https://contoso.sharepoint.com/sites/Build/Shared%20Documents/MyFiles/EU_01_MyFiles_Trade_And_CBAM_Tariff_Briefing.docx",
                        "extracts": [{"text": "EUR 124.8 billion bilateral goods & services turnover under CBAM."}],
                        "resourceMetadata": {
                            "title": "Europe Bilateral Trade & CBAM",
                            "RefinableString100": "Europe",
                        },
                    }
                ]
            }

        result = await retrieve_two_level_sharepoint(
            "What are the CBAM tariff rules?",
            "Europe",
            lambda: _token("delegated-token"),
            OPTIONS,
            post_request,
        )
        self.assertEqual(result.status, RetrievalStatus.SUCCESS)
        self.assertEqual(result.country, "Europe")
        self.assertEqual(len(calls), 2)
        self.assertIn('RefinableString100:"Europe"', calls[0]["filterExpression"])
        self.assertIn("Path:", calls[1]["filterExpression"])
        self.assertEqual(calls[1]["queryString"], "What are the CBAM tariff rules?")
        self.assertTrue(all(item.country == "Europe" for item in result.items))

    async def test_conversation_requires_country_selection_before_retrieval_and_filters_by_country(self):
        clear_conversation_states()
        messages: list[str] = []
        cards: list[dict] = []

        # Step 1: User asks a question before selecting a country
        res1 = await handle_two_level_conversation_turn(
            conversation_id="conv-1",
            text="What are the trade tariff policies and AI governance rules?",
            value=None,
            get_access_token=lambda: _token("delegated-token"),
            send_text=lambda t: _append(messages, t),
            send_card=lambda c: _append(cards, c),
            options=OPTIONS,
        )
        self.assertIsNone(res1)
        self.assertIn("Country", messages[-1])
        self.assertEqual(len(cards), 1)

        # Step 2: User selects 'US' -> runs Two-Level Retrieval on the pending question for US only
        for target_country, expected_prefix in (("US", "US_"), ("Europe", "EU_"), ("China", "CN_")):
            cands = get_country_candidate_files(target_country, OPTIONS.sharepoint_site_url)
            expected_count = 12 if target_country == "Europe" else 4
            self.assertEqual(len(cands), expected_count)
            self.assertTrue(all(c.filename.startswith(expected_prefix) for c in cands))
            self.assertTrue(all(c.country == target_country for c in cands))


async def _token(value: str) -> str:
    return value


async def _append(items: list, value):
    items.append(value)


async def _response(status: int, payload: dict):
    return status, payload
