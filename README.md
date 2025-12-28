# Automated Billing Verifier

## Executive Summary
This project is a high-precision Verification Engine developed via Google Apps Script (GAS). 
It is engineered to safeguard data integrity between dynamic Subscription Contracts and the Customer Success Management (CSM) database. 
By resolving cross-file data retrieval bottlenecks and implementing multi-layered verification logic, the engine proactively detects billing discrepancies, ensuring 100% accuracy before billing cycle execution.

## Project Structure
1. [interface.js](./interface.js) - UI initialization and custom menu management.
2. [checkTPR.js](./checkTPR.js) - Triple Check engine: Cross-verifies contract details with historical CSM records.
3. [ReportDialog.html](./ReportDialog.html) - Interactive HTML5 sidebar for discrepancy visualization and navigation.
4. [checkCB.js](./checkCB.js) - Billing & Plan Auditor: Validates billing rows against master Plan tables (RSV/Deposit).

## UI Preview & Interactive Features
* **Hierarchical View**: Each record is grouped by **Restaurant Name** (displayed as the header of each card for quick identification).
* **Status Indicators**: Uses color-coded borders (Yellow/Red) to highlight discrepancies.
* **Direct Navigation**: Clicking any card will trigger `jumpToRow()` to locate the data in the spreadsheet.
* *(Note: Restaurant names are hidden in the preview images below due to privacy policies.)*

<p align="center">
  <img src="./screenshots/discrepancy_sidebar.png" height="400" alt="Discrepancy Detected" />
  &nbsp;&nbsp;&nbsp;&nbsp;
  <img src="./screenshots/success_sidebar.png" height="400" alt="All Consistent" />
</p>
<p align="center">
  <em>Left: Discrepancy Alert Mode | Right: Success Confirmation Mode</em>
</p>

## Core Technical Contributions
### Multi-Faceted Validation Logic
* **Shared ID Integrity**: Utilizes a ```SharedIdMap``` to synchronize overage pools across multiple restaurant entities.
* **Deposit & Commission Logic**: Automatically distinguishes between B2B and B2C structures based on fee presence, enforcing strict format and numerical fidelity.
* **Diverse Output Methods**: Combines Dynamic Background Matrix (Sheet highlighting) with an Interactive HTML Sidebar (Data itemization) for comprehensive auditing.
### Performance Optimization & Latency Mitigation
* **Sectional Data Indexing**: Replaced iterative API calls with a Sectional Retrieval strategy, fetching specific data blocks (ID & Billing) into memory.
* **Startup Acceleration**: Migrated from ```openById``` to ```getActiveSpreadsheet``` for local operations, reducing unstable execution time from 250+ seconds to under 20 seconds.
### Setup and Deployment
This project utilizes modern version control practices and is deployed via Clasp.
* **Prerequisites**: Node.js, npm, and Clasp.
* **Secure Configuration**: Database IDs are managed via Script Properties (```PropertiesService```) to ensure zero exposure in source code.
