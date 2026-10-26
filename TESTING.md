---
title: Two-Level Retrieval Demo Testing Guide
description: Step-by-step guide to verify dummy files, SharePoint Country Classification metadata populated by the SharePoint agent, unit tests, and the Web Chat Two-Level Retrieval flow.
ms.date: 2026-10-26
ms.topic: how-to
---

## Overview

Use this guide to test and demonstrate the Two-Level Retrieval workflow with the Microsoft 365 Copilot Retrieval API across `MyFiles` and `OrgFiles` on SharePoint `site2` (`https://m365cpi50688407.sharepoint.com/sites/site2`).

The SharePoint library uses the custom column `Country Classification`, not a column named `Country`. The SharePoint agent has already populated this column with the country or region each document belongs to. The supplied screenshots show `China`, `Europe`, and `US` values in both `Documents > MyFiles` and `Documents > OrgFiles`.

The local manifests and generated documents still use the metadata key `Country`, and the Web Chat cards still use the label **Country**. These refer to the same classification values; the SharePoint column to verify and map is `Country Classification`.

The testing flow covers five stages:

1. Verify the local 20-document corpus (`Sharepoint/MyFiles` and `Sharepoint/OrgFiles`) and its `Country` metadata values.
2. Verify the uploaded files and the SharePoint-agent-populated `Country Classification` values on SharePoint `site2`.
3. Configure or verify the mapping from the custom column's crawled property to `RefinableString100` in SharePoint Search Schema.
4. Run the automated unit test suite.
5. Execute the live conversational demo in Azure Bot Web Chat.

## Step 1: Verify the Local Dummy Files and Metadata Manifests

Confirm that all 20 six-page `.docx` files (6 `Europe`, 2 `US`, and 2 `China` per folder) and the metadata manifests exist under `Sharepoint/`.

1. Open a terminal in the repository root and activate the virtual environment:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

2. Regenerate or verify the 20 `.docx` files and metadata manifests:

   ```powershell
   python scripts\generate_demo_files.py
   ```

3. Inspect the generated files in `Sharepoint/MyFiles` and `Sharepoint/OrgFiles`:

   | Folder                | Europe (6 files)        | US (2 files)            | China (2 files)         | Total    |
   | --------------------- | ----------------------- | ----------------------- | ----------------------- | -------- |
   | `Sharepoint/MyFiles`  | `EU_01` through `EU_06` | `US_01` through `US_02` | `CN_01` through `CN_02` | 10 files |
   | `Sharepoint/OrgFiles` | `EU_01` through `EU_06` | `US_01` through `US_02` | `CN_01` through `CN_02` | 10 files |

   Use each file's local `Country` value as the expected SharePoint `Country Classification` value.

4. Open any generated `.docx` file in Microsoft Word (for example `Sharepoint/MyFiles/EU_01_MyFiles_Trade_And_CBAM_Tariff_Briefing.docx`) to confirm:
   * Page 1 displays the classification table with `Country (SharePoint Column)` and `Managed Property (RefinableString100)`.
   * The document contains 6 full pages of country-specific analysis covering trade, tariffs, clean energy, critical minerals, AI governance, security, and ministerial summit talking points.
   * File > Info > Properties > Advanced Properties > Custom shows `Country`, `RefinableString100`, `RepositoryFolder`, and `DocumentId`.

## Step 2: Verify SharePoint Files and Country Classification Values

The files are already uploaded to SharePoint `site2`, and the SharePoint agent has updated `Country Classification` to match each document's country or region. Verify the existing values rather than creating a separate `Country` column or repeating the upload.

### Verify the Existing SharePoint-Agent-Populated Values

1. Open both SharePoint folders in your browser:
   * [MyFiles on site2](https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgDoY09VZbkjSbhtdzGiIxDLAWqgVZdBXeYMVkMw3ugS0cM?e=QCWbRD)
   * [OrgFiles on site2](https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgBiVQP4gpXzToCBCpbdqgtTAVA8za08GM8R6wQbeMDa9EE?e=4fkr7v)
2. Confirm that the existing `Country Classification` column is visible in both folder views, as shown in the supplied screenshots.
3. Compare its values with the `Country` values in `Sharepoint/metadata_manifest.csv`:

   | File prefix | Country Classification | MyFiles | OrgFiles | Total |
   | ----------- | ---------------------- | ------- | -------- | ----- |
   | `EU_`       | `Europe`               | 6       | 6        | 12    |
   | `US_`       | `US`                   | 2       | 2        | 4     |
   | `CN_`       | `China`                | 2       | 2        | 4     |
   | All files   | No blank values        | 10      | 10       | 20    |

4. Spot-check each classification against the document's content. Use the exact values `Europe`, `US`, and `China`; do not substitute `EU`, `USA`, or `United States`.

### Recreate the Setup Only If Needed

Skip this subsection for the already-populated library.

1. Upload the 10 `.docx` files from `Sharepoint/MyFiles` into the SharePoint `MyFiles` folder and the 10 files from `Sharepoint/OrgFiles` into `OrgFiles`.
2. If configuring a new library, create the custom column `Country Classification`. Reuse the existing column if it is already present.
3. Ask the SharePoint agent to populate the existing column based on each document's country or region. For example:

   ```text
   For every document in MyFiles and OrgFiles, update the existing Country Classification column to match the country or region the document belongs to. Use Europe for EU_ files, US for US_ files, and China for CN_ files. Check the document content for consistency. Do not create a separate Country column.
   ```

4. Review the populated values and counts against the table above before proceeding.

> [!WARNING]
> The current [publisher script](./scripts/publish_to_sharepoint.py) and the Web Chat `sync sharepoint` command
> use the [publisher service](./src/services/sharepoint_publisher.py), which creates or writes the legacy `Country` field.
> They do not populate `Country Classification`. Use the SharePoint agent or edit the existing column in the browser
> for this setup; do not treat the publisher's success message as confirmation that the custom column was updated.

## Step 3: Verify `RefinableString100` Search Schema Mapping

To enable server-side KQL filtering, map the crawled property for `Country Classification` to the existing managed property `RefinableString100`. The KQL property name remains `RefinableString100`; it does not change to the column's display name.

1. Open **Documents** > **Library settings** > **Columns** > **Country Classification**. Note the column's internal field name from the `Field` parameter in the settings page URL. The display name alone does not establish the internal name, which may contain encoded spaces or reflect an earlier name.
2. In SharePoint `site2`, go to **Settings** (gear icon) > **Site information** > **View all site settings** > **Search Schema** (under Site Collection Administration).
3. Under **Crawled Properties**, locate the property generated from that internal field name. Select the actual property for `Country Classification`; do not assume it is `ows_Country`. If the new column's crawled property is not yet available, request a library reindex as described below, wait for a crawl, and check again.
4. Under **Managed Properties**, find `RefinableString100` and map it to the crawled property identified in the previous step, then save. If an earlier demo mapping points to a separate legacy `Country` field, replace that mapping for this setup so filtering uses `Country Classification`.
5. Go to **Documents** > **Library settings** > **Advanced settings** and select **Reindex Document Library** so the search crawler indexes the custom column values into `RefinableString100`.
6. After indexing, verify the live Level 1 results for each selection:
   * `RefinableString100:"Europe"`: 12 demo files (6 per folder)
   * `RefinableString100:"US"`: 4 demo files (2 per folder)
   * `RefinableString100:"China"`: 4 demo files (2 per folder)

   Retain the `site2` path restriction shown in Scenario A. If live results are missing or incorrect, recheck the custom column values, actual crawled property mapping, and indexing status.

> [!NOTE]
> SharePoint search indexing can take 15 to 60 minutes or longer after metadata updates or schema mapping.
> While SharePoint indexing is in progress, the agent still executes the live Level 1 and Level 2 calls against
> `POST https://graph.microsoft.com/v1.0/copilot/retrieval` using your delegated Graph token and automatically grounds
> from the Level 1-approved local `Sharepoint/MyFiles` and `Sharepoint/OrgFiles` `.docx` files so you can test and demo immediately.
> Local fallback results do not verify that `Country Classification` has been indexed or mapped correctly;
> confirm that the expected files are returned by live retrieval before marking the SharePoint metadata test complete.

## Step 4: Run the Automated Unit Tests

Run the unit test suite to verify the 20 dummy `.docx` files (page counts, custom XML properties, and country counts), Level 1 and Level 2 KQL `filterExpression` builders, Two-Level Retrieval API requests, and the conversational Country selection state machine:

```powershell
python -m unittest discover -s tests -v
```

Expected output: all 12 unit tests pass (`OK`).

These tests validate the local `Country` metadata and retrieval logic, not the live SharePoint `Country Classification` values or search schema mapping. Complete Steps 2 and 3 to verify those separately.

## Step 5: Test the Conversational Experience in Web Chat

### 5.1 Start the Agent Server and Dev Tunnel

1. Ensure the Python agent server is running on port `3978`:

   ```powershell
   python -m src.main
   ```

2. In a second terminal, ensure the Dev Tunnel is hosting port `3978`:

   ```powershell
   devtunnel host -p 3978 --allow-anonymous
   ```

3. In the Azure Portal, open **SharePointDemo-AzureBot** > **Test in Web Chat** and complete the OAuth sign-in prompt if requested.

### 5.2 Scenario A: Ask a Question Before Selecting a Country

This scenario verifies that the agent enforces Country selection before retrieving any content.

1. In Web Chat, type the following question before choosing a country:

   ```text
   What are the key trade agreement priorities, tariff policies, and AI governance regulations?
   ```

2. Verify the agent response:
   * The agent does not run unfiltered retrieval.
   * It queues your question and displays the **Two-Level Retrieval: Select Country Filter** Adaptive Card with buttons for **Europe (12 files)**, **US (4 files)**, and **China (4 files)**.
3. Select **Europe (12 files)** on the card (or type `Europe`).
4. Verify the Two-Level Retrieval output for **Europe**:
   * **Level 1 (4.1)** shows the KQL filter `path:"https://m365cpi50688407.sharepoint.com/sites/site2/" AND RefinableString100:"Europe"` and lists only the 12 `Europe` files (6 in `MyFiles`, 6 in `OrgFiles`).
   * **Level 2 (4.2)** runs your semantic question strictly over those 12 `Europe` file paths (`(Path:".../EU_01_MyFiles_..." OR ...)`).
   * The returned extracts and source card contain only **Europe** content (CBAM carbon border tariffs, CETA, EU AI Act, DG TRADE) from `MyFiles` and `OrgFiles`, with each source tagged `Country: Europe (RefinableString100="Europe")`.

### 5.3 Scenario B: Switch Country to `US` and Ask the Same Question

This scenario demonstrates to the client that changing the Country filter changes the candidate document pool and semantic answers for the exact same question.

1. Select **Switch Country** on the response card (or type `change country`).
2. Select **US (4 files)** on the Country Selection card (or type `US`).
3. Verify the **Level 1 Filter Active: Country: US** card:
   * Shows `RefinableString100:"US"`.
   * Lists only the 4 `US` files:
     * `MyFiles`: `US_01_MyFiles_USMCA_Trade_Tariffs_And_Supply_Chain_Briefing.docx`, `US_02_MyFiles_AI_Safety_Cybersecurity_And_Clean_Tech_Memo.docx`
     * `OrgFiles`: `US_01_OrgFiles_United_States_Strategic_Trade_And_Industrial_Policy_Dossier.docx`, `US_02_OrgFiles_United_States_Technology_AI_Energy_And_Security_Framework.docx`
4. Select the **Ask: Trade & Tariffs** button on the card (or paste the same question from Scenario A).
5. Verify the Two-Level Retrieval output for **US**:
   * Level 2 retrieves only from the 4 confirmed `US` files (`Country: US`, `RefinableString100="US"`), returning USMCA Joint Review, USTR Section 301/232 tariffs, US NIST AI RMF, and Inflation Reduction Act details with zero Europe or China files.

### 5.4 Scenario C: Switch Country to `China` and Ask the Same Question

1. Select **Switch Country** (or type `change country`), then select **China (4 files)** (or type `China`).
2. Verify the **Level 1 Filter Active: Country: China** card lists only the 4 `China` files (2 in `MyFiles`, 2 in `OrgFiles`) tagged with `Country=China` (`RefinableString100:"China"`).
3. Select **Ask: AI & Digital Rules** or **Ask: Clean Energy & Minerals** (or ask your own question).
4. Verify the Two-Level Retrieval output for **China**:
   * Level 2 retrieves only from the 4 confirmed `China` files (`Country: China`, `RefinableString100="China"`), returning CAC Generative AI Measures, PIPL cross-border data rules, MOFCOM export licensing for gallium/germanium/graphite, and 15th Five-Year Plan clean energy targets.

### 5.5 Scenario D: Inline Country Prefix Shortcut

You can also specify the country and question in a single message at any time:

```text
Europe: What are the critical raw materials and battery passport rules?
```

```text
US: Summarize the CHIPS Act semiconductor and NORAD modernization priorities.
```

```text
China: What are the MOFCOM export controls on graphite and rare earths?
```
