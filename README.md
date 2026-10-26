---
title: Build Genie Two-Level Retrieval Agent
description: Configure a SharePoint-grounded agent using Two-Level Retrieval with the Microsoft 365 Copilot Retrieval API and Country metadata filtering.
---

## Overview

Build Genie is a Microsoft 365 Agents SDK sample that demonstrates Two-Level Retrieval with the Microsoft 365 Copilot Retrieval API across `MyFiles` and `OrgFiles` in SharePoint. It uses the signed-in user's delegated Microsoft Graph token, so every result respects that user's SharePoint permissions.

## Two-Step (Two-Level) Retrieval Process

The Microsoft 365 Copilot Retrieval API (`POST https://graph.microsoft.com/v1.0/copilot/retrieval`) is stateless and does not return a server-side result-set handle. To guarantee that semantic answers are grounded exclusively in documents belonging to a user-selected jurisdiction (`Europe`, `US`, or `China`), this agent implements a deterministic **Two-Step (Two-Level) Retrieval** workflow across `src/agent.py`, `src/services/message_route.py`, `src/services/retrieval_client.py`, and `src/services/card.py`.

### Conversational Pre-Check: Country Selection Before Retrieval

Before any retrieval request runs, `handle_two_level_conversation_turn` in `src/services/message_route.py` verifies that the user has selected a target `Country` (`Europe`, `US`, or `China`):

1. When a user joins Web Chat (or types `change country`, `reset`, or `menu`), the agent displays the **Select Country Filter** Adaptive Card (`create_country_selection_card`).
2. If the user asks a natural-language question before choosing a country, the agent stores that question as `pending_question` in conversation state and prompts the user to pick `Europe`, `US`, or `China` first.
3. Once the user selects a country (or sends an inline prefix such as `Europe: What are the CBAM tariff rules?`), the agent immediately runs the two-step retrieval pipeline below.

### Step 1 (Level 1): Deterministic Country Metadata Filtering

In Step 1, `retrieve_two_level_sharepoint` in `src/services/retrieval_client.py` narrows the candidate document pool to files tagged with the chosen `Country` in `MyFiles` and `OrgFiles`:

1. `create_level1_filter_expression` combines the SharePoint site path filter with the `RefinableString100` managed property (mapped to the SharePoint `Country` column):

   ```text
   path:"https://m365cpi50688407.sharepoint.com/sites/site2/" AND RefinableString100:"Europe"
   ```

2. The agent sends the Level 1 discovery request to `POST https://graph.microsoft.com/v1.0/copilot/retrieval` using the signed-in user's delegated Microsoft Graph token:

   ```json
   {
     "queryString": "Find briefing source documents matching the selected country scope Europe",
     "dataSource": "oneDriveBusiness",
     "filterExpression": "path:\"https://m365cpi50688407.sharepoint.com/sites/site2/\" AND RefinableString100:\"Europe\"",
     "resourceMetadata": ["title", "author", "RefinableString100", "Country"],
     "maximumNumberOfResults": 25
   }
   ```

3. The agent extracts and confirms the eligible Country files (`level1_candidates`) across `MyFiles` and `OrgFiles`:
   * `Europe`: 12 eligible files (6 in `MyFiles`, 6 in `OrgFiles`)
   * `US`: 4 eligible files (2 in `MyFiles`, 2 in `OrgFiles`)
   * `China`: 4 eligible files (2 in `MyFiles`, 2 in `OrgFiles`)

### Step 2 (Level 2): Semantic Search Scoped Only to Confirmed Country Files

In Step 2, the agent runs the user's natural-language question strictly against the file paths confirmed in Step 1:

1. `create_level2_filter_expression` constructs an explicit KQL `Path` `OR` filter containing only the confirmed Level 1 Country document URLs:

   ```text
   (Path:"https://m365cpi50688407.sharepoint.com/sites/site2/Shared Documents/MyFiles/EU_01_MyFiles_Trade_And_CBAM_Tariff_Briefing.docx" OR Path:"https://m365cpi50688407.sharepoint.com/sites/site2/Shared Documents/OrgFiles/EU_01_OrgFiles_European_Union_Comprehensive_Trade_Policy_Framework.docx" OR ...)
   ```

2. The agent sends the Level 2 semantic query request to `POST https://graph.microsoft.com/v1.0/copilot/retrieval`:

   ```json
   {
     "queryString": "What are the key trade agreement priorities, tariff policies, and AI governance regulations?",
     "dataSource": "oneDriveBusiness",
     "filterExpression": "(Path:\"https://m365cpi50688407.sharepoint.com/sites/site2/Shared Documents/MyFiles/EU_01_...\" OR ...)",
     "resourceMetadata": ["title", "author", "RefinableString100", "Country"],
     "maximumNumberOfResults": 3
   }
   ```

3. The agent formats both stages in the Web Chat reply and renders an Adaptive Card (`create_source_card`) displaying:
   * The Level 1 KQL filter expression (`RefinableString100:"<Country>"`) and the complete list of confirmed Country files in `MyFiles` and `OrgFiles`.
   * The Level 2 semantic extracts, each badged with `Country`, `RefinableString100`, `Folder` (`MyFiles` or `OrgFiles`), `File`, and an `Open source` link.

### End-to-End Implementation Map

| File | Responsibility in Two-Step Retrieval |
| --- | --- |
| `src/agent.py` | Welcomes the user with the Country Selection card and routes incoming Web Chat messages and card actions with the delegated `GRAPH` token. |
| `src/services/message_route.py` | Maintains per-conversation country state (`ConversationRetrievalState`), enforces Country selection before retrieval, and formats the Level 1 and Level 2 verification output. |
| `src/services/retrieval_client.py` | Builds Level 1 (`RefinableString100`) and Level 2 (`Path` `OR`) KQL filters and executes both `POST /v1.0/copilot/retrieval` stages. |
| `src/services/card.py` | Builds the Country Selection card, Level 1 Scope preview card, and Level 2 Two-Level Retrieval verification source card. |
| `src/services/sharepoint_publisher.py` | Creates the `Country` column on SharePoint `site2`, uploads the 20 `.docx` files to `MyFiles` and `OrgFiles`, and tags each file's `Country` metadata. |

## Demo Corpus (`Sharepoint/MyFiles` and `Sharepoint/OrgFiles`)

The local `Sharepoint/` directory includes 20 six-page `.docx` briefing documents (10 in `Sharepoint/MyFiles` and 10 in `Sharepoint/OrgFiles`) along with `Sharepoint/metadata_manifest.json` and `Sharepoint/metadata_manifest.csv`. Each document contains OpenXML custom properties (`Country` and `RefinableString100`) and six pages of country-specific analysis covering trade tariffs, clean energy, critical minerals, AI governance, security, and ministerial summit talking points:

| Folder | Europe (`RefinableString100="Europe"`) | US (`RefinableString100="US"`) | China (`RefinableString100="China"`) | Total |
| --- | --- | --- | --- | --- |
| `Sharepoint/MyFiles` | 6 files (`EU_01` to `EU_06`) | 2 files (`US_01` to `US_02`) | 2 files (`CN_01` to `CN_02`) | 10 files |
| `Sharepoint/OrgFiles` | 6 files (`EU_01` to `EU_06`) | 2 files (`US_01` to `US_02`) | 2 files (`CN_01` to `CN_02`) | 10 files |

To regenerate the 20 `.docx` files and manifests locally:

```bash
python scripts/generate_demo_files.py
```

The current SharePoint setup uses the custom `Country Classification` column populated by the SharePoint agent.
Follow [TESTING.md](./TESTING.md) to verify its values and search schema mapping.
The legacy publisher below uploads the 20 files and creates or writes a separate `Country` field;
it does not populate `Country Classification`. Use it only for the legacy `Country`-column setup
(requires `Files.ReadWrite.All` and `Sites.Manage.All`):

```bash
python scripts/publish_to_sharepoint.py --interactive
```

> [!NOTE]
> The Retrieval API supports both `sharePoint` and `oneDriveBusiness` data sources. This sample uses `sharePoint` by default.
> If `sharePoint` returns no results for an indexed document, set `RETRIEVAL_DATA_SOURCE=oneDriveBusiness` and restart the application before changing permissions or reindexing.
> This alternative selector can return documents stored in a SharePoint library. It does not move the documents, and the same configured SharePoint site filter still applies.
> Use this alternative with a Microsoft 365 Copilot-licensed user. Retrieval API pay-as-you-go does not support the OneDrive data source.

## Prerequisites

* Python 3.10 or later
* An Azure subscription and an [Azure Bot](https://aka.ms/AgentsSDK-CreateBot)
* A Microsoft 365 tenant with Copilot Retrieval API entitlement, a user who can sign in, and a SharePoint site that user can read
* [Dev Tunnels](https://learn.microsoft.com/azure/developer/dev-tunnels/get-started) for Web Chat testing

Upload [ContosoBuildSessions2025.docx](Sharepoint/ContosoBuildSessions2025.docx) to the configured SharePoint site's **Documents** library.

## Configure Azure Bot OAuth

For local testing, create one single-tenant Microsoft Entra ID app registration to use for the Azure Bot and its OAuth connection.

1. In **Microsoft Entra ID** > **App registrations**, create a single-tenant app registration. Record its application (client) ID and directory (tenant) ID.
2. In **Certificates & secrets**, create a client secret and copy its value.
3. In **Authentication**, add the Web redirect URI `https://token.botframework.com/.auth/web/redirect`.
4. In **API permissions**, add Microsoft Graph delegated permissions `Files.Read.All` and `Sites.Read.All`, then grant admin consent.
5. On the Azure Bot resource, create an OAuth connection with these values:

   | Field | Value |
   | --- | --- |
   | Name | `graph` |
   | Service provider | Microsoft Entra ID v2 / Azure Active Directory v2 |
   | Client ID | The app registration application (client) ID |
   | Client secret | The app registration client-secret value |
   | Tenant ID | The app registration directory (tenant) ID |
   | Scopes | `Files.Read.All Sites.Read.All offline_access` |
   | Token Exchange URL | Leave blank |

6. Save the connection. Select **Test connection**, sign in as the test user, and confirm it succeeds.

## Configure the sample

Copy `env.TEMPLATE` to `.env`. Keep `.env` outside source control.

```powershell
Copy-Item env.TEMPLATE .env
```

```dotenv
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=<YOUR_BOT_APP_ID>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=<YOUR_BOT_APP_SECRET>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=<YOUR_TENANT_ID>
AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME=graph
RETRIEVAL_SHAREPOINT_SITE_URL=https://contoso.sharepoint.com/sites/Build
RETRIEVAL_DATA_SOURCE=sharePoint
RETRIEVAL_MAXIMUM_NUMBER_OF_RESULTS=3
```

`RETRIEVAL_SHAREPOINT_SITE_URL` must be the SharePoint site root. For a document at `https://contoso.sharepoint.com/sites/Build/Shared%20Documents/session.docx`, set `https://contoso.sharepoint.com/sites/Build`. Do not use a file URL or sharing link. `RETRIEVAL_MAXIMUM_NUMBER_OF_RESULTS` must be from `1` through `25`.

`RETRIEVAL_DATA_SOURCE` accepts only `sharePoint` or `oneDriveBusiness` and defaults to `sharePoint` when omitted. Both choices retain the configured site filter. Restart the Python application after changing these settings.

### Separate bot and SharePoint tenants

The bot identity and Graph user authorization can use separate app registrations.
If SharePoint belongs to a different tenant, register a Graph OAuth app in that tenant with the delegated permissions and redirect URI described above.
Use that app's client ID, secret, and tenant ID in the Azure Bot OAuth connection named `graph`.

Keep the Azure Bot's top-level Microsoft App ID and App Tenant ID, and the `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__*` values in `.env`, associated with the original bot app.
Sign in through the bot as a user entitled to use the Retrieval API and authorized to read the target SharePoint content.

## Run and test

1. Create an environment, install dependencies, and start the sample:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   python -m src.main
   ```

2. In a second terminal, expose it through a tunnel:

   ```powershell
   devtunnel user login
   devtunnel host -p 3978 --allow-anonymous
   ```

3. In Azure Bot, set the messaging endpoint to `{tunnel-url}/api/messages`.
4. Open **Test in Web Chat** and complete sign-in when prompted.
5. Select a **Country** (`Europe`, `US`, or `China`) on the Adaptive Card (or ask a question first and select the Country when prompted).
6. Ask a cross-cutting question such as `What are the key trade agreement priorities, tariff policies, and AI governance regulations?` and verify that Level 1 lists only the selected Country's files in `MyFiles` and `OrgFiles` and Level 2 returns extracts exclusively from those Country files.

For a complete step-by-step walkthrough of all test scenarios, see [TESTING.md](./TESTING.md).

If restarting the tunnel produces a different public URL, update the Azure Bot messaging endpoint before testing again.

## Run unit tests

```powershell
python -m unittest discover -s tests -p test_retrieval.py -v
```

## Wait for SharePoint indexing

The document can appear in the Documents library before SharePoint search and Retrieval API can find it. Indexing can take minutes or hours; SharePoint does not provide a per-file indexing status.

Before testing the agent, sign in as the same Web Chat user and search the site for the document title or a unique phrase such as `Pricing Analytics`. Test the agent only after SharePoint search returns the document.

If the document is still not searchable after a reasonable wait, first confirm that the library permits search results. In the current SharePoint UI, go to **Site contents**, select the **three dots** for **Documents**, then select **Settings** > **Advanced settings**. Under **Search**, set **Allow items from this document library to appear in search results** to **Yes**. Select **Reindex Document Library** on the same page if needed. Reindexing adds the library to the next crawl; it does not complete immediately. Request it once only. See [Microsoft's reindex guidance](https://learn.microsoft.com/en-us/sharepoint/crawl-site-content#reindex-a-site).

## Troubleshooting

| Symptom | Action |
| --- | --- |
| Startup fails with `RETRIEVAL_SHAREPOINT_SITE_URL` | Set an absolute HTTPS SharePoint site-root URL in `.env`. |
| Sign-in does not complete | Confirm the Azure Bot OAuth connection name is `graph`, then check its app registration and delegated permissions. |
| No results | Confirm the user can open the document, the configured URL is the site root, and SharePoint site search finds the document. If needed, request one library reindex. |
| Retrieval is unavailable | Confirm tenant entitlement, Microsoft Graph permissions, and service availability; retry later. |

## Security notes

* The application never writes delegated tokens, Microsoft Graph response bodies, or stack traces to chat.
* Microsoft Graph applies the signed-in user's SharePoint permissions to Retrieval API requests.
* Keep client secrets in `.env`, environment variables, or a managed secret store.

## Further reading

* [Microsoft 365 Agents SDK](https://learn.microsoft.com/microsoft-365/agents-sdk/)
* [Microsoft 365 Copilot Retrieval API](https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval)
