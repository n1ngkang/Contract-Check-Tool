# Contract and Billing Verifier

## Executive Summary
This GAS-powered Verification Engine safeguards financial integrity through a dual-layer audit system: the checkTPR engine ensures CSM database synchronization, 
while the checkCB module automates complex RSV/Deposit billing validation.

## Project Structure
1. [interface.js](./interface.js) - UI initialization and custom menu management.
2. [checkTPR.js](./checkTPR.js) - Triple Check engine: Cross-verifies contract details with historical CSM records.
3. [ReportDialog.html](./ReportDialog.html) - Interactive HTML5 sidebar for discrepancy visualization and navigation.
4. [checkCB.js](./checkCB.js) - Billing & Plan Auditor: Validates billing rows against master Plan tables (RSV/Deposit).

## UI Preview & Interactive Features
* **Hierarchical View**: Each record is grouped by **Restaurant Name** (displayed as the header of each card for quick identification).
* **Dynamic Status Indicators**: Utilizes color-coded visual cues (e.g., Yellow/Red borders) to instantly flag data discrepancies and urgent issues.
* **Direct Contract Access**: Each record now includes a direct Contract Link. If a discrepancy is found, users can instantly jump to the source document to verify the original data, eliminating the need to manual search.
* *(Note: Restaurant names are hidden in the preview images below due to privacy policies.)*

<p align="center">
  <img src="./screenshots/discrepancy_sidebar.png" height="400" alt="Discrepancy Detected" />
  &nbsp;&nbsp;&nbsp;&nbsp;
  <img src="./screenshots/success_sidebar.png" height="400" alt="All Consistent" />
</p>
<p align="center">
  <em>Left: Discrepancy Alert Mode | Right: Success Confirmation Mode</em>
</p>

## Core Engine Logic
* **Data Integrity Sync (`checkTPR`)**: Ensures the "Source of Truth" by triple-checking active contracts against **CSM database** snapshots, preventing unauthorized data decay or manual input errors.
* **Precision Billing Audit (`checkCB`)**: Validates complex **RSV/Deposit** schemas, utilizing a `SharedIdMap` for overage consistency and **Price Range Evaluation** to distinguish between standard and negotiated plans.
* **Diverse Output Strategy**: Combines **Atomic Sheet Highlighting** (for bulk visual cues) with an **Interactive HTML5 Sidebar** (for granular discrepancy navigation).

## Technical Highlights
* **Performance Optimization**: Implemented **Sectional Memory Caching** to reduce cross-file latency, slashing execution time from **250s to <20s**.
* **Enterprise-Grade Security**: Sensitive database IDs and API endpoints are managed via **Script Properties** (`PropertiesService`), ensuring zero credential exposure in source code.
