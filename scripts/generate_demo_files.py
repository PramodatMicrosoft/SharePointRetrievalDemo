# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

"""Generate 20 multi-page (6-page) .docx briefing documents across MyFiles and OrgFiles
with Country metadata (6 Europe, 2 US, 2 China per folder) and RefinableString100 mapping.
"""

from __future__ import annotations

import csv
import json
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

SITE_URL = "https://m365cpi50688407.sharepoint.com/sites/site2"
MYFILES_SHARE_LINK = "https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgDoY09VZbkjSbhtdzGiIxDLAWqgVZdBXeYMVkMw3ugS0cM?e=QCWbRD"
ORGFILES_SHARE_LINK = "https://m365cpi50688407.sharepoint.com/:f:/s/site2/IgBiVQP4gpXzToCBCpbdqgtTAVA8za08GM8R6wQbeMDa9EE?e=4fkr7v"


@dataclass(frozen=True)
class DocSpec:
    filename: str
    folder: str  # "MyFiles" or "OrgFiles"
    country: str  # "Europe", "US", or "China"
    doc_id: str
    title: str
    subtitle: str
    author: str
    diplomatic_mission: str
    key_institutions: str
    flagship_policy: str
    trade_metrics: str
    energy_program: str
    digital_framework: str
    security_posture: str
    supply_chain_focus: str
    summit_deliverables: str


DOC_SPECS: list[DocSpec] = [
    # ==================== MyFiles: 6 Europe, 2 US, 2 China ====================
    DocSpec(
        filename="EU_01_MyFiles_Trade_And_CBAM_Tariff_Briefing.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-001",
        title="Europe Bilateral Trade & CBAM Carbon Border Tariff Assessment",
        subtitle="Analyst Working Draft — European Commission DG TRADE & Customs Adjustment Mechanism",
        author="Elena Rostova (Desk Officer, Europe Trade Division)",
        diplomatic_mission="Mission of Canada to the European Union (Brussels)",
        key_institutions="European Commission (DG TRADE, DG TAXUD), European Parliament INTA Committee, European Central Bank",
        flagship_policy="EU Carbon Border Adjustment Mechanism (CBAM) Regulation (EU) 2023/956 & CETA Provisional Application",
        trade_metrics="EUR 124.8 billion bilateral goods & services turnover; CBAM transitional reporting covering steel, aluminum, cement, fertilizers, electricity, and hydrogen imports into the European single market",
        energy_program="REPowerEU Cross-Border Hydrogen Corridor & North Sea Offshore Grid Interconnect",
        digital_framework="EU Customs Data Hub Digital Reporting & Single Window Certificates",
        security_posture="European Economic Security Strategy joint outbound investment screening",
        supply_chain_focus="Low-carbon green steel verification, Rotterdam and Antwerp port customs clearance, and European Authorised CBAM Declarant registries",
        summit_deliverables="Joint CETA Bilateral Trade Committee protocol on mutual recognition of carbon pricing and SME tariff facilitation in Brussels",
    ),
    DocSpec(
        filename="EU_02_MyFiles_AI_Act_And_Digital_Sovereignty_Memo.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-002",
        title="Europe AI Act Enforcement, GDPR & Digital Markets Briefing Memo",
        subtitle="Analyst Working Draft — European Artificial Intelligence Office & Digital Sovereignty",
        author="Marc-Antoine Lefebvre (Senior Digital Policy Analyst, Europe Desk)",
        diplomatic_mission="Embassy of Paris & EU Digital Attaché Network (Brussels / Luxembourg)",
        key_institutions="European AI Office, European Data Protection Board (EDPB), ENISA, DG CONNECT",
        flagship_policy="EU Artificial Intelligence Act (Regulation (EU) 2024/1689), Digital Markets Act (DMA), and Digital Services Act (DSA)",
        trade_metrics="EUR 48.2 billion cross-border ICT and cloud services trade across the 27 EU Member States; administrative fines up to 7% of global turnover for prohibited AI practices",
        energy_program="Green Cloud Data Centre Efficiency Directive under the European Energy Efficiency Framework",
        digital_framework="High-Risk AI System Conformity Assessments, General-Purpose AI (GPAI) systemic risk codes of practice, EuroHPC AI Factories (JUPITER & LEONARDO supercomputers), and Gaia-X federated data spaces",
        security_posture="NIS2 Directive critical infrastructure incident reporting and ENISA European Cybersecurity Certification Scheme for Cloud Services (EUCS)",
        supply_chain_focus="European sovereign cloud infrastructure, Franco-German algorithmic transparency audits, and GDPR Adequacy Decision renewal",
        summit_deliverables="EU-Canada Digital Partnership Ministerial Council joint declaration on interoperable AI safety testing and cross-Atlantic research sandboxes",
    ),
    DocSpec(
        filename="EU_03_MyFiles_Clean_Energy_Green_Deal_And_Hydrogen_Note.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-003",
        title="Europe Green Deal, Net-Zero Industry Act & Green Hydrogen Corridor Note",
        subtitle="Analyst Working Draft — European Climate Law & Trans-European Energy Infrastructure",
        author="Hannah Lindqvist (Clean Energy & Climate Attaché, Europe Bureau)",
        diplomatic_mission="Embassy in Berlin & Nordic-Baltic Energy Coordination Cell",
        key_institutions="DG ENER, DG CLIMA, European Hydrogen Bank, European Investment Bank (EIB)",
        flagship_policy="European Green Deal, Fit for 55 Package, Net-Zero Industry Act (NZIA), and Renewable Energy Directive III (RED III)",
        trade_metrics="EUR 300 billion REPowerEU investment mobilization; European Hydrogen Bank auction premiums of EUR 0.48/kg for RFNBO certified green ammonia and hydrogen imports via Hamburg and Wilhelmshaven",
        energy_program="42.5% binding EU renewable energy target by 2030, Baltic and North Sea 120 GW offshore wind expansion, and H2Med pipeline",
        digital_framework="European Smart Grids Digital Twin Action Plan and automated Guarantee of Origin (GoO) ledger",
        security_posture="Protection of European subsea energy pipelines, LNG terminal resilience in Germany and Netherlands, and phase-out of Russian fossil fuel imports",
        supply_chain_focus="Electrolyzer manufacturing capacity in Europe (40% domestic benchmark under NZIA) and green shipping fuel corridors",
        summit_deliverables="Signing of the Hamburg-Rotterdam Atlantic Clean Hydrogen Offtake Annex and EIB co-financing facility",
    ),
    DocSpec(
        filename="EU_04_MyFiles_NATO_Security_And_Defense_Industrial_Brief.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-004",
        title="Europe Security Architecture, European Defence Industrial Strategy & PESCO Brief",
        subtitle="Analyst Working Draft — European External Action Service (EEAS) & Transatlantic Pillar",
        author="Col. (Ret.) Jean-François Moreau (European Security & Defence Analyst)",
        diplomatic_mission="Permanent Delegation in Brussels & Warsaw Security Liaison Office",
        key_institutions="European External Action Service (EEAS), European Defence Agency (EDA), NATO HQ Brussels, Frontex",
        flagship_policy="European Defence Industrial Strategy (EDIS), European Defence Investment Programme (EDIP), and Strategic Compass for Security and Defence",
        trade_metrics="EUR 100 billion SAFE European defence procurement facility; target of 50% intra-European defence equipment collaborative procurement by 2030",
        energy_program="Military Mobility Green Fuel Standardization across the Trans-European Transport Network (TEN-T)",
        digital_framework="EU Cyber Solidarity Act, European Cyber Shield Security Operations Centres, and IRIS2 secure satellite constellation",
        security_posture="Eastern Flank deterrence readiness in Poland and the Baltic states, Permanent Structured Cooperation (PESCO) military mobility corridors, and hybrid threat countering",
        supply_chain_focus="155mm artillery ammunition ramp-up under ASAP (Act in Support of Ammunition Production), European air defense (Sky Shield), and dual-use component licensing",
        summit_deliverables="Bilateral Security & Defence Partnership implementation roadmap with EEAS in Brussels",
    ),
    DocSpec(
        filename="EU_05_MyFiles_Critical_Raw_Materials_And_Battery_Supply_Chain.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-005",
        title="Europe Critical Raw Materials Act (CRMA) & Battery Passport Supply Chain Memo",
        subtitle="Analyst Working Draft — European Battery Alliance & Strategic Mineral Partnerships",
        author="Dr. Sofia Rossi (Industrial Policy & Minerals Specialist, Europe Division)",
        diplomatic_mission="Consulate General in Milan & European Industrial Liaison Office",
        key_institutions="DG GROW, European Raw Materials Alliance (ERMA), EIT InnoEnergy",
        flagship_policy="EU Critical Raw Materials Act (Regulation (EU) 2024/1252) and EU Battery Regulation (EU) 2023/1542 Digital Battery Passport",
        trade_metrics="EU 2030 statutory benchmarks: 10% domestic extraction, 40% European processing, 25% recycling of 17 strategic raw materials (lithium, nickel, cobalt, rare earths, battery-grade graphite)",
        energy_program="Low-carbon hydroelectric and nuclear refining recognition for European gigafactory cathode precursors in Sweden, France, and Germany",
        digital_framework="Mandatory QR-code EU Digital Battery Passport tracking carbon footprint, recycled content, and ESG due diligence from mine to European EV assembly",
        security_posture="Strategic Diversification Cap: no more than 65% of any strategic raw material consumed in the European Union sourced from a single third country",
        supply_chain_focus="Joint European-allied strategic mineral extraction projects, ESG traceability under the EU Corporate Sustainability Due Diligence Directive (CSDDD)",
        summit_deliverables="Designation of 5 joint transatlantic mining and refining projects as EU Strategic Projects eligible for accelerated permitting and EIB financing",
    ),
    DocSpec(
        filename="EU_06_MyFiles_Ministerial_Summit_Talking_Points_Brussels.docx",
        folder="MyFiles",
        country="Europe",
        doc_id="GAC-MY-EU-006",
        title="Europe Ministerial Summit Aide-Mémoire & Talking Points (Brussels / Strasbourg)",
        subtitle="Analyst Working Draft — Bilateral Ministerial Preparation & Horizon Europe",
        author="Claire Beauchamp (Chief Briefing Coordinator, Europe Bilateral Relations)",
        diplomatic_mission="Mission to the European Union, Avenue de Cortenbergh, Brussels",
        key_institutions="European Council, European Commission College of Commissioners, Horizon Europe Pillar II Directorate",
        flagship_policy="Strategic Partnership Agreement (SPA), Horizon Europe Association Pillar II, and Schengen ETIAS Mobility",
        trade_metrics="EUR 68.4 billion bidirectional foreign direct investment stock across Germany, France, Netherlands, Ireland, and Italy; 99% tariff line elimination under CETA",
        energy_program="Trans-Atlantic Clean Aviation SAF (Sustainable Aviation Fuel) corridor between European hubs (Frankfurt, Amsterdam, Paris CDG) and North America",
        digital_framework="Horizon Europe Pillar II digital, quantum, and photonics joint calls (EUR 95.5 billion total programme budget)",
        security_posture="Coordinated G7 and European sanctions enforcement against illicit maritime shadow fleets in the Baltic and Mediterranean seas",
        supply_chain_focus="Pharmaceutical active ingredient resilience under the European Critical Medicines Alliance and agricultural GI (Geographical Indications) protections",
        summit_deliverables="Adoption of the Brussels Joint Ministerial Communiqué covering CETA ratification progress, Horizon Europe research grants, and Arctic-European scientific cooperation",
    ),
    DocSpec(
        filename="US_01_MyFiles_USMCA_Trade_Tariffs_And_Supply_Chain_Briefing.docx",
        folder="MyFiles",
        country="US",
        doc_id="GAC-MY-US-001",
        title="US Bilateral Trade, USMCA Joint Review & Section 301 Tariff Assessment",
        subtitle="Analyst Working Draft — USTR, Department of Commerce & North American Supply Chains",
        author="David Harrington (Senior Trade Policy Officer, United States Bureau)",
        diplomatic_mission="Embassy in Washington, D.C. (Pennsylvania Avenue NW)",
        key_institutions="Office of the United States Trade Representative (USTR), US Department of Commerce (ITA & BIS), US International Trade Commission (USITC), CBP",
        flagship_policy="United States-Mexico-Canada Agreement (USMCA / CUSMA) Article 34.7 Joint Review, Buy America Act (BAA), and Section 301 / Section 232 Trade Remedies",
        trade_metrics="USD 960 billion annual bilateral goods and services trade ($2.6 billion daily cross-border commerce across Detroit-Windsor, Buffalo-Fort Erie, and Pacific Highway corridors)",
        energy_program="North American Cross-Border Electricity Transmission & Keystone/Enbridge Energy Security Corridor",
        digital_framework="USMCA Chapter 19 Digital Trade prohibitions on customs duties for digital products and CBP ACE Single Window automation",
        security_posture="CFIUS (Committee on Foreign Investment in the United States) outbound and inbound investment screening and Defense Production Act (DPA) Title III",
        supply_chain_focus="USMCA 75% Regional Value Content (RVC) automotive rules of origin, softwood lumber countervailing duty litigation, and agricultural dairy TRQ administration",
        summit_deliverables="Washington Free Trade Commission (FTC) Joint Statement on USMCA 2026 review readiness, border pre-clearance expansion, and critical supply chain early-warning mechanism",
    ),
    DocSpec(
        filename="US_02_MyFiles_AI_Safety_Cybersecurity_And_Clean_Tech_Memo.docx",
        folder="MyFiles",
        country="US",
        doc_id="GAC-MY-US-002",
        title="US AI Safety Institute, CHIPS Act & Inflation Reduction Act Clean Tech Memo",
        subtitle="Analyst Working Draft — US NIST AI RMF, CISA Cybersecurity & DOE Clean Energy Credits",
        author="Dr. Priya Nair (Science, Technology & Energy Counsellor, US Desk)",
        diplomatic_mission="Embassy in Washington, D.C. & Consulates in San Francisco and Boston",
        key_institutions="US NIST (AI Safety Institute), CISA (Cybersecurity and Infrastructure Security Agency), US Department of Energy (DOE), DARPA",
        flagship_policy="US Executive Order on Safe, Secure, and Trustworthy AI, NIST AI Risk Management Framework (AI RMF 1.0), CHIPS and Science Act, and Inflation Reduction Act (IRA Section 30D & 45X)",
        trade_metrics="USD 52.7 billion CHIPS and Science Act semiconductor incentives and USD 369 billion Inflation Reduction Act clean energy tax credits for North American assembly",
        energy_program="US DOE Loan Programs Office (LPO) grid-scale storage, Small Modular Reactor (SMR) HALEU fuel supply, and North American NERC grid reliability standards",
        digital_framework="US AISI Frontier Model Red-Teaming Protocols, CISA Secure by Design software bill of materials (SBOM) mandates, and FedRAMP High cloud reciprocity",
        security_posture="NORAD modernization radar architecture, CISA Shields Up critical infrastructure defense against Volt Typhoon, and BIS Export Administration Regulations (EAR)",
        supply_chain_focus="IRA Section 30D Foreign Entity of Concern (FEOC) exclusion compliance for EV battery components and critical minerals processed in North America",
        summit_deliverables="Bilateral Cross-Border AI Safety Testing Memorandum with US NIST and joint DOE Grid Resilience Task Force workplan",
    ),
    DocSpec(
        filename="CN_01_MyFiles_Bilateral_Trade_Export_Controls_And_Tariffs_Briefing.docx",
        folder="MyFiles",
        country="China",
        doc_id="GAC-MY-CN-001",
        title="China Bilateral Trade, MOFCOM Export Licensing & Agricultural Tariff Briefing",
        subtitle="Analyst Working Draft — MOFCOM, GACC Customs & Strategic Export Controls",
        author="Kevin Zhang (Senior Trade & Economic Analyst, China & Northeast Asia Division)",
        diplomatic_mission="Embassy in Beijing (Chaoyang District) & Consulate General in Shanghai",
        key_institutions="Ministry of Commerce of the People's Republic of China (MOFCOM), General Administration of Customs (GACC), NDRC, State Council Tariff Commission",
        flagship_policy="PRC Export Control Law, MOFCOM Dual-Use Item Export Licensing Catalogue (Gallium, Germanium, Antimony, Graphite), and GACC Decree 248/249 Food Facility Registration",
        trade_metrics="RMB 745 billion (USD 104.2 billion) bilateral merchandise trade; canola seed, potash, wood pulp, and iron ore exports alongside EV, solar cell, and lithium-ion battery tariff investigations",
        energy_program="China National Energy Administration (NEA) Ultra-High Voltage (UHV) transmission and petrochemical LNG import terminals in Guangdong and Zhejiang",
        digital_framework="China International Trade Single Window (GACC) electronic phytosanitary e-Cert verification",
        security_posture="PRC Counter-Espionage Law compliance for foreign chambers of commerce and MOFCOM Unreliable Entity List monitoring",
        supply_chain_focus="MOFCOM end-user and end-use certificates for synthetic and natural spherical graphite, rare earth permanent magnets, and port congestion at Ningbo-Zhoushan and Shenzhen",
        summit_deliverables="Convening of the Joint Economic and Trade Commission (JETC) working group on GACC agricultural inspection timelines and maritime shipping lanes",
    ),
    DocSpec(
        filename="CN_02_MyFiles_AI_Governance_EV_Batteries_And_Clean_Energy_Memo.docx",
        folder="MyFiles",
        country="China",
        doc_id="GAC-MY-CN-002",
        title="China Generative AI Measures, PIPL Cross-Border Data & EV Clean Tech Memo",
        subtitle="Analyst Working Draft — Cyberspace Administration of China (CAC) & MIIT Industrial Policy",
        author="Dr. Mei-Ling Lau (Digital & Industrial Technology Officer, China Desk)",
        diplomatic_mission="Embassy in Beijing & Consulates in Guangzhou and Chongqing",
        key_institutions="Cyberspace Administration of China (CAC), Ministry of Industry and Information Technology (MIIT), National Data Bureau (NDB), Ministry of Ecology and Environment (MEE)",
        flagship_policy="CAC Interim Measures for the Management of Generative Artificial Intelligence Services, Personal Information Protection Law (PIPL), Data Security Law (DSL), and 15th Five-Year Plan New Quality Productive Forces",
        trade_metrics="Over 60% global share of EV LFP/NMC battery cell manufacturing and 80% solar PV wafer capacity; RMB 1.8 trillion clean-tech investment across Yangtze and Pearl River Deltas",
        energy_program="China 1,200+ GW wind and solar installed capacity milestone, National ETS (Emissions Trading System) expansion to aluminum and cement, and Sinopec green hydrogen pipelines",
        digital_framework="CAC Algorithm Filing Registry, Large Language Model security assessments, and CAC Provisions on Promoting and Regulating Cross-Border Data Flows (Pilot Free Trade Zone negative lists in Shanghai, Beijing, and Hainan)",
        security_posture="Multi-Level Protection Scheme (MLPS 2.0) and Critical Information Infrastructure Operator (CIIO) local data storage rules in mainland China",
        supply_chain_focus="CATL and BYD battery supply chain vertical integration, MIIT lithium battery industry norms, and multinational R&D localization in Shanghai and Shenzhen",
        summit_deliverables="Technical dialogue with China MEE on Kunming-Montreal Global Biodiversity Framework implementation and methane abatement monitoring",
    ),

    # ==================== OrgFiles: 6 Europe, 2 US, 2 China ====================
    DocSpec(
        filename="EU_01_OrgFiles_European_Union_Comprehensive_Trade_Policy_Framework.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-101",
        title="Europe Institutional Policy Framework: CETA Trade & Single Market Regulatory Alignment",
        subtitle="Departmental Reference Dossier — European Union Trade, Customs & Regulatory Cooperation",
        author="Directorate-General for European Trade & Economic Policy (OrgFiles Repository)",
        diplomatic_mission="Global Affairs Departmental Headquarters & EU Directorate (Brussels / Ottawa)",
        key_institutions="European Commission DG TRADE, European Council Trade Policy Committee, European Court of Justice, CEN/CENELEC",
        flagship_policy="Comprehensive Economic and Trade Agreement (CETA), EU Carbon Border Adjustment Mechanism (CBAM) Definitive Regime (2026), and EU Deforestation Regulation (EUDR)",
        trade_metrics="EUR 124.8 billion total bilateral trade; 98.6% of EU tariff lines duty-free; EUR 14.5 billion services trade surplus in financial, engineering, and aviation sectors across Europe",
        energy_program="EU-Allied Clean Fuel Standard alignment and European Maritime FuelEU carbon intensity thresholds",
        digital_framework="EU EIDAS 2.0 European Digital Identity Wallet and Mutual Recognition of Conformity Assessment (CETA Protocol on Conformity Assessment)",
        security_posture="EU Anti-Coercion Instrument (Regulation (EU) 2023/2675) and Foreign Subsidies Regulation (FSR) M&A notification thresholds",
        supply_chain_focus="EUDR geolocation traceability for timber, cattle, soy, and paper entering European Union ports (Rotterdam, Hamburg, Antwerp, Marseille)",
        summit_deliverables="Institutional 5-Year CETA Regulatory Cooperation Forum Workplan and Joint Customs Cooperation Committee Decision",
    ),
    DocSpec(
        filename="EU_02_OrgFiles_European_Digital_Markets_And_AI_Compliance_Standard.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-102",
        title="Europe Institutional Standard: EU AI Act, Data Act & Digital Sovereignty Governance",
        subtitle="Departmental Reference Dossier — European Digital Single Market & Algorithmic Governance",
        author="Office of the Chief Digital & International Cyber Policy Officer (OrgFiles Repository)",
        diplomatic_mission="European Union Digital Policy Coordination Secretariat (Brussels / Tallinn / Paris)",
        key_institutions="European AI Board, European Centre for Algorithmic Transparency (ECAT - Seville), EDPB, BEREC",
        flagship_policy="EU AI Act (2024/1689), EU Data Act (Regulation (EU) 2023/2854), Cyber Resilience Act (CRA), and GDPR Chapter V International Transfers",
        trade_metrics="EUR 195 billion European cloud and AI market valuation; CRA mandatory CE-marking cybersecurity lifecycle rules for all hardware and software placed on the European market",
        energy_program="European Commission Code of Conduct for Data Centre Sustainability and waste-heat recovery mandates in Germany and Denmark",
        digital_framework="EU Data Act cloud switching interoperability (eliminating egress fees across European providers), EuroHPC quantum computing hubs, and European Health Data Space (EHDS)",
        security_posture="Cyber Resilience Act 24-hour severe vulnerability notification to ENISA and CSIRT networks across all 27 EU Member States",
        supply_chain_focus="CE marking for high-risk AI components in medical devices, aviation, and automotive machinery sold in the European Economic Area (EEA)",
        summit_deliverables="Ratification of the Administrative Arrangement on Artificial Intelligence Safety Evaluation and GDPR Cross-Border Data Adequacy Review",
    ),
    DocSpec(
        filename="EU_03_OrgFiles_European_NetZero_Energy_Transition_And_Climate_Pact.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-103",
        title="Europe Institutional Dossier: European Climate Law, ETS2 & Net-Zero Energy Transition",
        subtitle="Departmental Reference Dossier — European Decarbonization, Hydrogen & Offshore Grids",
        author="International Climate & Energy Diplomacy Bureau (OrgFiles Repository)",
        diplomatic_mission="European Climate & Energy Coordination Hub (Brussels / Copenhagen / Berlin)",
        key_institutions="European Environment Agency (EEA - Copenhagen), DG ENER, ACER (Agency for the Cooperation of Energy Regulators), EIB",
        flagship_policy="European Climate Law (Legally Binding 2050 Climate Neutrality & 2040 90% Emission Reduction Target), EU Emissions Trading System (ETS & ETS2), and Net-Zero Industry Act",
        trade_metrics="EUR 85/tonne average EU ETS carbon allowance price; EUR 65 billion Social Climate Fund and EUR 40 billion Innovation Fund for European industrial decarbonization",
        energy_program="North Seas Energy Cooperation (NSEC) 300 GW offshore wind blueprint by 2050, European Hydrogen Backbone (53,000 km pipeline network across 28 European countries), and PCI (Projects of Common Interest)",
        digital_framework="ACER REMIT II algorithmic wholesale energy trading surveillance and European Grid Digitalization Action Plan",
        security_posture="AggregateEU joint gas purchasing mechanism and mandatory 90% European underground gas storage filling targets prior to winter",
        supply_chain_focus="European wind turbine manufacturing charter, high-voltage direct current (HVDC) subsea cable production in Italy, France, and Sweden, and carbon capture and storage (CCS) in the North Sea",
        summit_deliverables="High-Level Green Alliance Ministerial Declaration with the European Commission on carbon market linkage and transatlantic clean hydrogen certification",
    ),
    DocSpec(
        filename="EU_04_OrgFiles_European_Security_Architecture_And_Resilience_Dossier.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-104",
        title="Europe Institutional Dossier: Common Foreign & Security Policy (CFSP) & Resilience",
        subtitle="Departmental Reference Dossier — European Defence Union, Hybrid Deterrence & Critical Infrastructure",
        author="Strategic Policy & European Security Directorate (OrgFiles Repository)",
        diplomatic_mission="Transatlantic & European Security Policy Division (Brussels / Mons / The Hague)",
        key_institutions="Council of the European Union (Foreign Affairs Council), EEAS Crisis Management Directorate, Europol, Hybrid CoE (Helsinki)",
        flagship_policy="EU Common Security and Defence Policy (CSDP), Critical Entities Resilience (CER) Directive (EU) 2022/2557, and EU-NATO Joint Declaration III",
        trade_metrics="EUR 326 billion combined EU Member State defence expenditure (1.9% of EU GDP); EUR 8 billion European Defence Fund (EDF) collaborative R&D budget",
        energy_program="European Defence Energy Autonomous Camp and Synthetic Aviation Fuel defence logistics initiative",
        digital_framework="EU Space Strategy for Security and Defence, Galileo Public Regulated Service (PRS), and Copernicus Earth observation security services",
        security_posture="CER Directive resilience assessments across 11 critical European sectors (energy, transport, banking, health, drinking water, digital infrastructure, space) and EU Rapid Deployment Capacity (5,000 troops)",
        supply_chain_focus="Schengen Military Mobility corridor upgrades (bridges, rail gauges, customs Form 302 digitization) from Dutch and German ports to Poland, Romania, and the Baltic states",
        summit_deliverables="Signing of the Administrative Arrangement for Participation in European Defence Agency (EDA) Category B Projects and PESCO Military Mobility",
    ),
    DocSpec(
        filename="EU_05_OrgFiles_European_Semiconductor_Chips_Act_And_Critical_Minerals.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-105",
        title="Europe Institutional Dossier: European Chips Act & Strategic Raw Materials Resilience",
        subtitle="Departmental Reference Dossier — European Semiconductor Ecosystem & Mineral Processing",
        author="Advanced Manufacturing & Critical Supply Chains Bureau (OrgFiles Repository)",
        diplomatic_mission="European Industrial & Innovation Network (Leuven / Eindhoven / Dresden / Grenoble)",
        key_institutions="European Semiconductor Board, Chips Joint Undertaking (Chips JU), IMEC (Belgium), Fraunhofer (Germany), CEA-Leti (France)",
        flagship_policy="European Chips Act (Regulation (EU) 2023/1781) and Critical Raw Materials Act Strategic Projects Framework",
        trade_metrics="EUR 43 billion public-private European Chips Act investment targeting 20% of global semiconductor manufacturing value share in Europe by 2030",
        energy_program="Ultra-pure water and 100% renewable power purchase agreements (PPAs) for European megafabs in Dresden (Silicon Saxony), Magdeburg, and Crolles",
        digital_framework="European Cloud-Edge-IoT continuum, RISC-V open-source processor architecture consortium, and extreme ultraviolet (EUV) lithography ecosystem led by ASML in Eindhoven",
        security_posture="Pillar III Semiconductor Crisis Response Mechanism (mandatory supply-chain monitoring, priority rated orders) and coordinated European export controls on advanced semiconductor manufacturing equipment and quantum cryogenics",
        supply_chain_focus="Automotive microcontroller supply for European OEMs, silicon carbide (SiC) power electronics for EV inverters, and rare-earth refining in Estonia and France",
        summit_deliverables="Launch of the Europe-North America Semiconductor Supply Chain Early Warning Mechanism and Joint Research Call with IMEC and CEA-Leti",
    ),
    DocSpec(
        filename="EU_06_OrgFiles_European_Macroeconomic_Outlook_And_ECB_Fiscal_Report.docx",
        folder="OrgFiles",
        country="Europe",
        doc_id="GAC-ORG-EU-106",
        title="Europe Institutional Report: Eurozone Macroeconomic Outlook, ECB Policy & Capital Markets",
        subtitle="Departmental Reference Dossier — European Central Bank, NextGenerationEU & Fiscal Governance",
        author="Office of the Chief Economist — Europe & Eurozone Analysis Unit (OrgFiles Repository)",
        diplomatic_mission="Financial & Economic Counsellor Network (Frankfurt / Brussels / Luxembourg)",
        key_institutions="European Central Bank (ECB - Frankfurt), Eurogroup, European Stability Mechanism (ESM), European Securities and Markets Authority (ESMA)",
        flagship_policy="Reformed EU Economic Governance Framework (Stability and Growth Pact debt sustainability trajectories), NextGenerationEU (EUR 806.9 billion Recovery and Resilience Facility), and Capital Markets Union (Savings and Investments Union)",
        trade_metrics="EUR 15.8 trillion Eurozone GDP; 1.3% projected real GDP growth across the 20 Euro-area states; EUR 650 billion NextGenerationEU green and digital bond issuance",
        energy_program="ECB climate stress-testing of European banking balance sheets and European Green Bond Standard (EuGB) taxonomy alignment",
        digital_framework="ECB Digital Euro preparation phase (offline privacy, pan-European retail settlement, and DLT wholesale central bank money trials)",
        security_posture="European Anti-Money Laundering Authority (AMLA - Frankfurt) establishment and G7 Immobilized Sovereign Assets windfall revenue facility",
        supply_chain_focus="Draghi Report on European Competitiveness recommendations: closing the EUR 800 billion annual investment gap in European innovation, energy grids, and defense",
        summit_deliverables="Bilateral Finance Ministries Macroeconomic Dialogue in Frankfurt on sustainable finance taxonomies, Digital Euro interoperability, and sovereign debt resilience",
    ),
    DocSpec(
        filename="US_01_OrgFiles_United_States_Strategic_Trade_And_Industrial_Policy_Dossier.docx",
        folder="OrgFiles",
        country="US",
        doc_id="GAC-ORG-US-101",
        title="US Institutional Policy Dossier: USMCA Architecture, Industrial Policy & Trade Enforcement",
        subtitle="Departmental Reference Dossier — United States Trade, Customs, Supply Chains & Regulatory Policy",
        author="North American Trade & Economic Policy Directorate (OrgFiles Repository)",
        diplomatic_mission="Washington Bilateral Coordination Secretariat & 12 US Consulate Economic Posts",
        key_institutions="USTR, US Department of Commerce, US Department of the Treasury, Congressional Ways and Means & Senate Finance Committees",
        flagship_policy="USMCA (CUSMA) 2026 Six-Year Joint Review Mandate, Trade Act of 1974 Section 301, Trade Expansion Act Section 232, and Infrastructure Investment and Jobs Act (IIJA)",
        trade_metrics="USD 960 billion total US-Canada bilateral trade; 36 US states count Canada as their #1 export destination; USD 1.1 trillion bilateral direct investment relationship",
        energy_program="North American Energy Ministers Trilateral Framework, 35+ cross-border electricity interconnectors, and US Strategic Petroleum Reserve coordination",
        digital_framework="US-Canada Beyond the Border Automated Commercial Environment (ACE) and Trusted Trader C-TPAT / PIP harmonization",
        security_posture="Executive Order 14105 Outbound Investment Security Program (Treasury rules covering semiconductors, quantum information technologies, and AI systems) and Defense Production Act Title III co-investment",
        supply_chain_focus="North American EV and automotive supply-chain integration across Michigan, Ohio, Indiana, and Ontario; critical mineral co-processing under the Energy Resource Governance Initiative",
        summit_deliverables="Adoption of the Joint North American Competitiveness & Economic Security Action Plan ahead of the 2026 USMCA Joint Commission Review",
    ),
    DocSpec(
        filename="US_02_OrgFiles_United_States_Technology_AI_Energy_And_Security_Framework.docx",
        folder="OrgFiles",
        country="US",
        doc_id="GAC-ORG-US-102",
        title="US Institutional Dossier: National AI Strategy, CHIPS Act, Clean Energy & Continental Defense",
        subtitle="Departmental Reference Dossier — US Science, Technology, Clean Energy & Homeland Security",
        author="Continental Defence, Cyber & Emerging Technology Directorate (OrgFiles Repository)",
        diplomatic_mission="Washington Defence, Cyber & Science Policy Coordination Cell",
        key_institutions="White House OSTP, US Department of Defense (DoD), US Department of Energy (DOE), DHS / CISA, National Science Foundation (NSF)",
        flagship_policy="US National Security Memorandum on AI, CHIPS and Science Act (USD 280 billion authorization), Inflation Reduction Act Clean Energy Tax Architecture, and NORAD Modernization Plan",
        trade_metrics="USD 39 billion CHIPS manufacturing grants across Arizona, Ohio, Texas, and New York fabs; USD 38.6 billion bi-national NORAD continental defense investment",
        energy_program="US DOE Clean Hydrogen Hubs (H2Hubs - USD 7 billion across 7 regional US hubs), FERC Order 1920 long-term transmission planning, and fusion energy milestone",
        digital_framework="NIST Post-Quantum Cryptography (PQC) FIPS 203/204/205 migration standards, National AI Research Resource (NAIRR), and CISA Cross-Sector Cybersecurity Performance Goals",
        security_posture="Over-the-Horizon Arctic Radar system deployment, Five Eyes (FVEY) cyber threat intelligence sharing, and US Space Force Combined Space Operations Initiative",
        supply_chain_focus="Advanced semiconductor packaging corridor (Northeast Microelectronics Coalition & US-allied packaging facilities) and domestic uranium/HALEU nuclear fuel cycle",
        summit_deliverables="Joint Continental Technology & Defence Accord covering PQC cryptographic migration, AI safety evaluation standards, and Arctic domain awareness",
    ),
    DocSpec(
        filename="CN_01_OrgFiles_China_Economic_Trade_And_Supply_Chain_Policy_Dossier.docx",
        folder="OrgFiles",
        country="China",
        doc_id="GAC-ORG-CN-101",
        title="China Institutional Policy Dossier: Macroeconomic Dual Circulation, Trade & Supply Chains",
        subtitle="Departmental Reference Dossier — PRC Economic Governance, MOFCOM Trade & Industrial Policy",
        author="Indo-Pacific & China Economic Policy Directorate (OrgFiles Repository)",
        diplomatic_mission="Greater China Economic & Trade Analysis Network (Beijing / Shanghai / Hong Kong)",
        key_institutions="State Council of the PRC, National Development and Reform Commission (NDRC), MOFCOM, People's Bank of China (PBOC), GACC",
        flagship_policy="PRC Dual Circulation Strategy, 15th Five-Year Plan Industrial Modernization, Foreign Investment Law (24-Point State Council Guidelines), and Export Control Regulations on Dual-Use Items",
        trade_metrics="USD 17.9 trillion PRC nominal GDP (~4.8% annual growth target); RMB 3.2 trillion high-tech manufacturing lending facility by PBOC; USD 104 billion bilateral goods trade",
        energy_program="NDRC Dual-Control transition from energy intensity to carbon emissions intensity caps and 80 GW nuclear power operational/under-construction program across coastal provinces",
        digital_framework="PBOC e-CNY Digital Yuan cross-border mBridge wholesale settlement trials and GACC Smart Customs 2026 inspection platform",
        security_posture="PRC Anti-Foreign Sanctions Law, Law on Foreign Relations of the PRC, and national security reviews for foreign investments in strategic sectors",
        supply_chain_focus="China's dominant refining share in rare earths (85%), battery-grade lithium chemicals (65%), cobalt refining (75%), and polysilicon (80%), alongside WTO dispute settlement filings",
        summit_deliverables="Departmental Risk-Managed Engagement Framework for China Bilateral Trade, Agricultural Market Access Stability, and Supply Chain De-Risking",
    ),
    DocSpec(
        filename="CN_02_OrgFiles_China_Digital_Regulation_Clean_Tech_And_Diplomatic_Framework.docx",
        folder="OrgFiles",
        country="China",
        doc_id="GAC-ORG-CN-102",
        title="China Institutional Dossier: Cyber Sovereignty, AI Plus Initiative & Climate Diplomacy",
        subtitle="Departmental Reference Dossier — PRC Digital Governance, Green Manufacturing & Diplomatic Engagement",
        author="Northeast Asia Strategic & Regulatory Affairs Bureau (OrgFiles Repository)",
        diplomatic_mission="Beijing Diplomatic & Regulatory Monitoring Secretariat",
        key_institutions="Cyberspace Administration of China (CAC), Ministry of Foreign Affairs (MOFA), Ministry of Science and Technology (MOST), Ministry of Ecology and Environment (MEE)",
        flagship_policy="State Council 'AI Plus' (AI+) Initiative, PRC Cybersecurity Law / Data Security Law / PIPL Tri-Pillar Data Regime, and China 2030 Carbon Peak / 2060 Carbon Neutrality ('30-60' Dual Carbon Goals)",
        trade_metrics="Over 200 CAC-registered generative AI models in mainland China; 9.5 million annual New Energy Vehicle (NEV) domestic production and global export surge",
        energy_program="China National Carbon Market (world's largest ETS by covered emissions at 5.1 billion tonnes CO2) and Desert Gobi荒漠 (Shagehuang) 450 GW wind-solar megabase buildout",
        digital_framework="National Data Bureau (NDB) 'Data Elements X' three-year action plan, Eastern Data / Western Computing (Dongshu Xisuan) 8 national computing hubs, and CAC Standard Contract (SCC) for data exports",
        security_posture="Global AI Governance Initiative (GAIGI) promoted by MOFA, strict ICP licensing and content provenance watermarking mandates under CAC Deep Synthesis Regulations",
        supply_chain_focus="Closed-loop battery recycling rules under MIIT, green electricity certificates (GECs) for export manufacturers in Jiangsu and Guangdong, and maritime port automation",
        summit_deliverables="Bilateral Diplomatic & Environmental Engagement Protocol covering COP Climate Cooperation, Biodiversity Conservation, and Consular/Commercial Predictability",
    ),
]


def _build_six_pages_of_sections(spec: DocSpec) -> list[tuple[str, list[str]]]:
    """Return 6 rich page sections (1 page per section with explicit page break)
    totaling ~2,200+ words so every document exceeds 5 full pages in Word and clearly
    demonstrates Country-specific filtering in both Level 1 and Level 2 retrieval.
    """
    folder_type = (
        "Analyst Working Briefing & Operational Memo (MyFiles Personal/Team Workspace)"
        if spec.folder == "MyFiles"
        else "Official Departmental Policy Framework & Institutional Dossier (OrgFiles Enterprise Library)"
    )

    page1_paragraphs = [
        f"DOCUMENT CLASSIFICATION METADATA — COUNTRY: {spec.country} | MANAGED PROPERTY: RefinableString100={spec.country} | REPOSITORY FOLDER: {spec.folder} | DOCUMENT ID: {spec.doc_id}",
        f"1.1 Executive Mandate and Jurisdictional Scope ({spec.country}): This {folder_type} provides an authoritative, multi-page assessment exclusively scoped to {spec.country}. Within the Two-Level Retrieval Architecture for Microsoft 365 Copilot Retrieval API, this document is governed under the SharePoint library column Country = '{spec.country}' (mapped in the SharePoint Search Schema to the queryable managed property RefinableString100 = '{spec.country}') and stored inside the '{spec.folder}' folder on SharePoint site2 ({SITE_URL}/Shared Documents/{spec.folder}/{spec.filename}).",
        f"1.2 Lead Diplomatic Mission and Institutional Counterparts in {spec.country}: Primary operational coordination is led by {spec.diplomatic_mission}, engaging directly with {spec.key_institutions}. All policy recommendations, briefing talking points, and risk assessments in this dossier apply strictly to the {spec.country} bilateral and multilateral theatre.",
        f"1.3 Flagship Regulatory and Policy Anchor ({spec.country}): The central policy instrument governing current engagement with {spec.country} is {spec.flagship_policy}. Senior officials and briefing analysts preparing ministerial aide-mémoires for {spec.country} must anchor all commercial, technological, and security assessments in the compliance milestones established by {spec.key_institutions}.",
        f"1.4 Summary of Strategic Indicators for {spec.country}: Bilateral commercial volume and strategic exposure are characterized by {spec.trade_metrics}. Across the 2026–2028 planning horizon, departmental priorities for {spec.country} center on safeguarding supply chain resilience, aligning clean energy and digital standards, and executing the deliverables defined for {spec.summit_deliverables}.",
        f"1.5 Two-Level Retrieval Verification Marker ({spec.country} / {spec.folder}): When a user selects '{spec.country}' in the conversational webchat interface, Level 1 deterministic filtering narrows the SharePoint candidate document set to files tagged with Country='{spec.country}' (RefinableString100:'{spec.country}'). Level 2 semantic retrieval then extracts grounded passages exclusively from the confirmed {spec.country} file paths, ensuring zero cross-contamination from other jurisdictions.",
        f"1.6 Departmental Governance and Quality Assurance ({spec.country}): In accordance with the AIde-Mémoire Architecture Design Document v2.0 (Section 8.1), all source materials inside {spec.folder} classified under Country='{spec.country}' undergo mandatory metadata verification prior to briefing synthesis. The delegated Microsoft Graph token enforces SharePoint item-level security trimming while preserving strict country-boundary isolation for {spec.country} throughout both retrieval stages.",
        f"1.7 Comparative Baseline and Analytical Scope ({spec.country}): Briefing officers consulting this {spec.country} document should cross-reference the statistical annexes maintained by {spec.key_institutions} and the operational cables transmitted by {spec.diplomatic_mission}. Key performance indicators tracked across pages 2 through 6 include tariff schedules, clean energy capacity additions, AI conformity deadlines, critical mineral refining shares, and ministerial summit deliverables specific to {spec.country}.",
    ]

    page2_paragraphs = [
        f"2.1 Bilateral Trade Architecture, Tariff Schedules, and Market Access in {spec.country}: Economic and commercial relations with {spec.country} are currently shaped by {spec.flagship_policy}. Quantitative trade tracking confirms {spec.trade_metrics}. Customs authorities and trade negotiators interacting with {spec.key_institutions} have prioritized predictable rules of origin, tariff rate quota transparency, and dispute avoidance mechanisms specific to {spec.country}.",
        f"2.2 Sectoral Tariff Exposure and Industrial Competitiveness ({spec.country}): In the {spec.country} market, industrial exporters face evolving regulatory compliance requirements tied to border adjustments, technical standards, and domestic industrial incentives. Specifically, {spec.supply_chain_focus} represents the primary operational bottleneck requiring coordinated intervention by {spec.diplomatic_mission}.",
        f"2.3 Customs Digitization and Border Facilitation ({spec.country}): Trade compliance workflows with {spec.country} increasingly rely on {spec.digital_framework}. Exporters and logistics operators routing cargo into {spec.country} must submit verified electronic declarations demonstrating adherence to {spec.country} statutory standards, carbon accounting methodologies, and supply-chain origin audits.",
        f"2.4 Foreign Direct Investment (FDI) and Outbound/Inbound Screening ({spec.country}): Capital flows between domestic enterprises and {spec.country} counterparties are subject to heightened scrutiny under {spec.security_posture}. Legal and commercial teams evaluating joint ventures, greenfield facilities, or technology licensing in {spec.country} must complete pre-transaction assessments covering national security review thresholds and subsidy disclosure rules enforced by {spec.key_institutions}.",
        f"2.5 Commercial Policy Recommendations for {spec.country}: (a) Establish a dedicated bilateral technical working group with {spec.key_institutions} to resolve customs and conformity assessment friction; (b) publish updated sector guidance for SMEs exporting to {spec.country} under {spec.flagship_policy}; and (c) advance the trade deliverables scheduled for {spec.summit_deliverables}.",
        f"2.6 Non-Tariff Barriers, Sanitary/Phytosanitary Standards, and Technical Regulations ({spec.country}): Beyond headline customs duties, market access into {spec.country} depends on conformity with product safety, environmental labeling, and sector-specific certification regimes administered by {spec.key_institutions}. Commercial attachés at {spec.diplomatic_mission} conduct bi-weekly regulatory horizon scanning to prevent shipment holds at primary {spec.country} ports of entry.",
        f"2.7 Bilateral Trade Dispute Resolution and Institutional Mechanisms ({spec.country}): Where divergent interpretations arise under {spec.flagship_policy}, structured consultations through the relevant bilateral trade committees provide an early-resolution channel prior to formal arbitral or multilateral proceedings. All evidentiary submissions for {spec.country} trade reviews are indexed under Country='{spec.country}' (RefinableString100='{spec.country}').",
    ]

    page3_paragraphs = [
        f"3.1 Clean Energy Transition, Decarbonization Policies, and Grid Infrastructure in {spec.country}: Energy security and net-zero industrial transformation form a cornerstone of engagement with {spec.country}. Under {spec.energy_program}, {spec.country} has accelerated capital deployment into renewable power generation, clean hydrogen infrastructure, and industrial electrification overseen by {spec.key_institutions}.",
        f"3.2 Carbon Pricing, Emissions Trading, and Clean Tech Subsidies ({spec.country}): Industrial decarbonization in {spec.country} is reinforced through statutory carbon pricing and targeted fiscal instruments linked to {spec.flagship_policy}. As documented in recent economic filings ({spec.trade_metrics}), energy-intensive manufacturers operating in or exporting to {spec.country} must account for embedded lifecycle emissions, verified clean power purchase agreements, and certified low-carbon fuel standards.",
        f"3.3 Critical Minerals, Battery Value Chains, and Refining Capacity ({spec.country}): Securing resilient upstream and midstream mineral supply chains is directly tied to {spec.supply_chain_focus}. Within the {spec.country} jurisdiction, policy measures prioritize domestic and allied processing capacity for lithium, nickel, cobalt, battery-grade graphite, and rare earth permanent magnets required for electric vehicles and grid storage.",
        f"3.4 Cross-Border Energy Corridors and Clean Fuel Certification ({spec.country}): Bilateral technical teams led by {spec.diplomatic_mission} are advancing mutual recognition of guarantee-of-origin certificates, carbon intensity accounting, and clean fuel logistics under {spec.energy_program}. Harmonizing verification protocols with {spec.key_institutions} prevents duplicative testing and unlocks long-term offtake financing.",
        f"3.5 Energy & Climate Action Plan for {spec.country}: Priority actions for the upcoming quarter include finalizing clean energy technical annexes under {spec.summit_deliverables}, aligning sustainable finance taxonomies with {spec.country} regulators, and expanding joint R&D on grid resilience and next-generation nuclear/storage technologies.",
        f"3.6 Industrial Electrification and Green Manufacturing Benchmarks ({spec.country}): Heavy industry transition plans across {spec.country} require deep integration between zero-emission baseload generation, grid-scale battery storage, and green hydrogen feedstocks. Project developers engaging with {spec.key_institutions} under {spec.energy_program} must satisfy stringent lifecycle greenhouse gas reduction thresholds and domestic content or diversification benchmarks.",
        f"3.7 Climate Finance, Green Bonds, and Transition Taxonomy Alignment ({spec.country}): Institutional investors financing decarbonization projects in {spec.country} rely on standardized climate disclosure and taxonomy frameworks overseen by {spec.key_institutions}. Aligning bilateral export credit and project finance guarantees with {spec.country} sustainability standards remains a key deliverable under {spec.summit_deliverables}.",
    ]

    page4_paragraphs = [
        f"4.1 Artificial Intelligence Governance, Digital Sovereignty, and Data Regulation in {spec.country}: The digital regulatory landscape in {spec.country} is defined by {spec.digital_framework} and enforced by {spec.key_institutions}. Organizations deploying foundation models, high-risk AI systems, or cloud infrastructure in {spec.country} must comply with mandatory risk classifications, algorithmic transparency obligations, and post-market monitoring.",
        f"4.2 Cross-Border Data Flows, Privacy Compliance, and Localization Rules ({spec.country}): Data governance under {spec.flagship_policy} establishes strict legal mechanisms for international data transfers involving {spec.country}. Enterprises operating across {spec.country} must maintain validated transfer instruments, privacy impact assessments, and cybersecurity safeguards aligned with {spec.security_posture}.",
        f"4.3 Semiconductor Ecosystems, High-Performance Computing, and Quantum Research ({spec.country}): Technological competitiveness in {spec.country} is anchored by strategic investments in advanced microelectronics, AI compute infrastructure, and quantum technologies ({spec.digital_framework}). Supply-chain coordination focusing on {spec.supply_chain_focus} is essential to mitigate fabrication bottlenecks and export control disruptions.",
        f"4.4 Cybersecurity Standards and Critical Infrastructure Protection ({spec.country}): Under {spec.security_posture}, {spec.country} authorities require rapid incident reporting, software bill of materials (SBOM) attestation, and hardware supply-chain integrity across telecommunications, financial services, energy grids, and government networks.",
        f"4.5 Digital & AI Policy Directives for {spec.country}: Analysts preparing briefing materials for {spec.country} should highlight regulatory interoperability opportunities between domestic AI safety institutes and {spec.key_institutions}, while tracking compliance deadlines under {spec.flagship_policy}.",
        f"4.6 Frontier AI Model Evaluation, Red-Teaming, and Algorithmic Auditing ({spec.country}): Technical standards bodies in {spec.country} ({spec.key_institutions}) have instituted structured evaluation protocols for general-purpose and frontier AI models covering systemic safety, cybersecurity resilience, synthetic content provenance, and copyright transparency under {spec.digital_framework}.",
        f"4.7 Digital Public Infrastructure, Cloud Sovereignty, and Quantum Readiness ({spec.country}): Public-sector cloud procurement and digital identity ecosystems in {spec.country} increasingly mandate cryptographic agility, post-quantum migration roadmaps, and operational resilience verified by {spec.key_institutions}. Bilateral science and technology cooperation is coordinated through {spec.diplomatic_mission}.",
    ]

    page5_paragraphs = [
        f"5.1 Strategic Security Posture, Economic Resilience, and Geopolitical Risk in {spec.country}: The broader security and resilience environment in {spec.country} is governed by {spec.security_posture}. Assessments coordinated through {spec.diplomatic_mission} emphasize the close linkage between national security, critical infrastructure protection, and economic de-risking across the {spec.country} theatre.",
        f"5.2 Critical Supply Chain Vulnerabilities and De-Risking Measures ({spec.country}): Detailed supply-chain mapping for {spec.country} identifies high-priority dependencies in {spec.supply_chain_focus}. In partnership with {spec.key_institutions}, contingency protocols have been established to monitor chokepoints, diversify supplier concentration, and maintain strategic buffer inventories.",
        f"5.3 Export Controls, Sanctions Enforcement, and Dual-Use Technology Screening ({spec.country}): Compliance with dual-use export control regimes and multilateral restrictive measures is a top operational priority when engaging with {spec.country}. Regulatory authorities ({spec.key_institutions}) enforce rigorous end-user verification, intangible technology transfer controls, and financial screening tied to {spec.flagship_policy}.",
        f"5.4 Infrastructure Resilience and Crisis Response Mechanisms ({spec.country}): Joint preparedness exercises and early-warning consultations in {spec.country} address maritime logistics security, subsea cable and pipeline protection, cyber threat intelligence sharing, and emergency supply-chain coordination under {spec.security_posture}.",
        f"5.5 Risk Mitigation Matrix for {spec.country}: Senior decision-makers reviewing {spec.country} files from the {spec.folder} repository must weigh macroeconomic exposure ({spec.trade_metrics}) against regulatory shifts in {spec.flagship_policy}, ensuring all bilateral commitments include verifiable safeguard clauses.",
        f"5.6 Research Security, Dual-Use Academic Collaboration, and IP Protection ({spec.country}): Universities, national laboratories, and corporate R&D centers collaborating with entities in {spec.country} are subject to research security guidelines covering sensitive technology domains ({spec.digital_framework}). Vetting procedures coordinated with {spec.key_institutions} protect intellectual property while preserving open scientific exchange where appropriate.",
        f"5.7 Maritime, Aviation, and Multimodal Logistics Security ({spec.country}): Freight corridors linking domestic ports and air cargo hubs with {spec.country} require continuous customs-security coordination, advance cargo manifest screening, and resilience planning around {spec.supply_chain_focus}.",
    ]

    page6_paragraphs = [
        f"6.1 Ministerial Summit Objectives, Diplomatic Talking Points, and Deliverables ({spec.country}): In preparation for upcoming high-level engagements led by {spec.diplomatic_mission}, the primary diplomatic objective for {spec.country} is {spec.summit_deliverables}. Briefing packages compiled by {spec.author} synthesize inputs from trade, energy, digital, and security divisions.",
        f"6.2 Core Ministerial Talking Points for {spec.country} Counterparts ({spec.key_institutions}): "
        f"(1) Reaffirm commitment to rules-based bilateral trade and transparent market access across {spec.country} ({spec.trade_metrics}); "
        f"(2) Advance practical regulatory alignment on {spec.flagship_policy} to minimize unintended barriers for compliant industries; "
        f"(3) Deepen clean energy and critical mineral collaboration through {spec.energy_program} and {spec.supply_chain_focus}; and "
        f"(4) Strengthen digital trust, AI safety interoperability, and infrastructure resilience under {spec.digital_framework} and {spec.security_posture}.",
        f"6.3 Interdepartmental Coordination and Stakeholder Consultations ({spec.country}): Domestic industry associations, provincial/regional representatives, and research institutions have been consulted regarding {spec.country} market conditions. Stakeholders strongly support the targeted outcomes outlined in {spec.summit_deliverables} and request timely guidance on {spec.flagship_policy}.",
        f"6.4 Implementation Timeline and Next Steps (2026–2027) for {spec.country}: "
        f"Q1–Q2: Technical working group sessions with {spec.key_institutions} on customs, AI governance, and clean energy standards. "
        f"Q3: Mid-term review of supply chain resilience metrics ({spec.supply_chain_focus}) and bilateral investment flows. "
        f"Q4: Ministerial summit execution and formal adoption of {spec.summit_deliverables}.",
        f"6.5 Watchlist Indicators and Trigger Events ({spec.country}): Desk officers at {spec.diplomatic_mission} will monitor legislative calendars, tariff schedule gazettes, and regulatory enforcement actions issued by {spec.key_institutions}. Any material change to {spec.flagship_policy} triggers an immediate Level 1 and Level 2 retrieval refresh in the AIde-Mémoire system.",
        f"6.6 Document Governance & Retrieval Attestation: Document ID {spec.doc_id} | Author: {spec.author} | Repository Folder: {spec.folder} | SharePoint Classification Column: Country = {spec.country} | Search Managed Property: RefinableString100 = {spec.country} | Verified 6-Page Briefing Dossier for Microsoft 365 Copilot Two-Level Retrieval Demo.",
    ]

    return [
        (f"Page 1 — Executive Summary & {spec.country} Jurisdictional Mandate", page1_paragraphs),
        (f"Page 2 — {spec.country} Bilateral Trade, Tariffs & Market Access Analysis", page2_paragraphs),
        (f"Page 3 — {spec.country} Clean Energy Transition & Critical Minerals Strategy", page3_paragraphs),
        (f"Page 4 — {spec.country} AI Governance, Digital Sovereignty & Cyber Regulation", page4_paragraphs),
        (f"Page 5 — {spec.country} Strategic Security, Supply Chains & Export Controls", page5_paragraphs),
        (f"Page 6 — {spec.country} Ministerial Summit Talking Points & 2026 Deliverables", page6_paragraphs),
    ]


def _build_document_xml(spec: DocSpec) -> str:
    sections = _build_six_pages_of_sections(spec)
    body_parts: list[str] = []

    # Document Header Title & Subtitle
    body_parts.append(
        f'<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
        f'<w:t>{escape(spec.title)}</w:t></w:r></w:p>'
    )
    body_parts.append(
        f'<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>'
        f'<w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr>'
        f'<w:t>{escape(spec.subtitle)}</w:t></w:r></w:p>'
    )

    # Metadata Classification Table on Page 1
    meta_rows = [
        ("Country (SharePoint Column)", spec.country),
        ("Managed Property (RefinableString100)", spec.country),
        ("Repository Folder", spec.folder),
        ("Document ID", spec.doc_id),
        ("Lead Author / Unit", spec.author),
        ("Diplomatic Mission", spec.diplomatic_mission),
        ("Key Counterpart Institutions", spec.key_institutions),
        ("Flagship Policy Instrument", spec.flagship_policy),
        ("SharePoint Target URL", f"{SITE_URL}/Shared Documents/{spec.folder}/{spec.filename}"),
    ]
    table_xml = ['<w:tbl><w:tblPr><w:tblBorders>'
                 '<w:top w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>'
                 '<w:left w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>'
                 '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>'
                 '<w:right w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>'
                 '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9E2F3"/>'
                 '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9E2F3"/>'
                 '</w:tblBorders></w:tblPr>']
    for k, v in meta_rows:
        table_xml.append(
            f'<w:tr>'
            f'<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{escape(k)}</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:p><w:r><w:t>{escape(v)}</w:t></w:r></w:p></w:tc>'
            f'</w:tr>'
        )
    table_xml.append('</w:tbl>')
    body_parts.append("".join(table_xml))

    for idx, (heading, paragraphs) in enumerate(sections):
        if idx > 0:
            # Explicit page break before Pages 2..6 to guarantee minimum 6 pages in Word
            body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        body_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>'
            f'<w:t>{escape(heading)}</w:t></w:r></w:p>'
        )
        for para in paragraphs:
            body_parts.append(
                f'<w:p><w:r><w:rPr><w:sz w:val="22"/></w:rPr>'
                f'<w:t xml:space="preserve">{escape(para)}</w:t></w:r></w:p>'
            )

    body_xml = "\n".join(body_parts)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body_xml}<w:sectPr/></w:body></w:document>"
    )


def _build_content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '<Override PartName="/docProps/custom.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
        "</Types>"
    )


def _build_root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        '<Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" '
        'Target="docProps/custom.xml"/>'
        "</Relationships>"
    )


def _build_core_xml(spec: DocSpec) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{escape(spec.title)}</dc:title>"
        f"<dc:subject>{escape(spec.subtitle)}</dc:subject>"
        f"<dc:creator>{escape(spec.author)}</dc:creator>"
        f"<cp:keywords>Country:{escape(spec.country)};RefinableString100:{escape(spec.country)};Folder:{escape(spec.folder)}</cp:keywords>"
        f"<dc:description>Country={escape(spec.country)} | RefinableString100={escape(spec.country)} | Folder={escape(spec.folder)} | {escape(spec.flagship_policy)}</dc:description>"
        f"<cp:category>{escape(spec.country)}</cp:category>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-20T09:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-24T10:00:00Z</dcterms:modified>'
        "</cp:coreProperties>"
    )


def _build_app_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Microsoft Office Word</Application>"
        "<Pages>6</Pages>"
        "<Company>Contoso Global Affairs</Company>"
        "</Properties>"
    )


def _build_custom_xml(spec: DocSpec) -> str:
    """OpenXML custom properties automatically promoted by SharePoint to list item columns."""
    fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f'<property fmtid="{fmtid}" pid="2" name="Country"><vt:lpwstr>{escape(spec.country)}</vt:lpwstr></property>'
        f'<property fmtid="{fmtid}" pid="3" name="RefinableString100"><vt:lpwstr>{escape(spec.country)}</vt:lpwstr></property>'
        f'<property fmtid="{fmtid}" pid="4" name="RepositoryFolder"><vt:lpwstr>{escape(spec.folder)}</vt:lpwstr></property>'
        f'<property fmtid="{fmtid}" pid="5" name="DocumentId"><vt:lpwstr>{escape(spec.doc_id)}</vt:lpwstr></property>'
        "</Properties>"
    )


def write_docx(path: Path, spec: DocSpec) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _build_content_types_xml())
        zf.writestr("_rels/.rels", _build_root_rels_xml())
        zf.writestr("word/document.xml", _build_document_xml(spec))
        zf.writestr("docProps/core.xml", _build_core_xml(spec))
        zf.writestr("docProps/app.xml", _build_app_xml())
        zf.writestr("docProps/custom.xml", _build_custom_xml(spec))


def generate_all(base_sharepoint_dir: Path | None = None) -> list[dict]:
    if base_sharepoint_dir is None:
        base_sharepoint_dir = Path(__file__).resolve().parent.parent / "Sharepoint"

    manifest_entries: list[dict] = []
    for spec in DOC_SPECS:
        out_path = base_sharepoint_dir / spec.folder / spec.filename
        write_docx(out_path, spec)
        share_folder_link = (
            MYFILES_SHARE_LINK if spec.folder == "MyFiles" else ORGFILES_SHARE_LINK
        )
        web_url = f"{SITE_URL}/Shared%20Documents/{spec.folder}/{spec.filename}"
        kql_path = f"{SITE_URL}/Shared Documents/{spec.folder}/{spec.filename}"
        entry = {
            "filename": spec.filename,
            "folder": spec.folder,
            "localPath": f"Sharepoint/{spec.folder}/{spec.filename}",
            "Country": spec.country,
            "RefinableString100": spec.country,
            "managedProperty": "RefinableString100",
            "docId": spec.doc_id,
            "title": spec.title,
            "subtitle": spec.subtitle,
            "author": spec.author,
            "pages": 6,
            "webUrl": web_url,
            "kqlPath": kql_path,
            "sharepointFolderLink": share_folder_link,
            "flagshipPolicy": spec.flagship_policy,
            "keyInstitutions": spec.key_institutions,
            "tradeMetrics": spec.trade_metrics,
            "energyProgram": spec.energy_program,
            "digitalFramework": spec.digital_framework,
            "securityPosture": spec.security_posture,
            "supplyChainFocus": spec.supply_chain_focus,
            "summitDeliverables": spec.summit_deliverables,
        }
        manifest_entries.append(entry)

    manifest_json_path = base_sharepoint_dir / "metadata_manifest.json"
    manifest_json_path.write_text(
        json.dumps(
            {
                "siteUrl": SITE_URL,
                "managedProperty": "RefinableString100",
                "columnName": "Country",
                "folders": {
                    "MyFiles": {
                        "shareLink": MYFILES_SHARE_LINK,
                        "counts": {"Europe": 6, "US": 2, "China": 2, "total": 10},
                    },
                    "OrgFiles": {
                        "shareLink": ORGFILES_SHARE_LINK,
                        "counts": {"Europe": 6, "US": 2, "China": 2, "total": 10},
                    },
                },
                "files": manifest_entries,
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    manifest_csv_path = base_sharepoint_dir / "metadata_manifest.csv"
    with manifest_csv_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "filename",
                "folder",
                "Country",
                "RefinableString100",
                "docId",
                "pages",
                "title",
                "author",
                "kqlPath",
                "webUrl",
            ],
        )
        writer.writeheader()
        for item in manifest_entries:
            writer.writerow({k: item[k] for k in writer.fieldnames})

    return manifest_entries


if __name__ == "__main__":
    entries = generate_all()
    print(f"Generated {len(entries)} .docx files across MyFiles and OrgFiles.")
