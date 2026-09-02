# Product Requirements Document (PRD)
## Outage EIC Work Order Monitoring System
**PLTU MSW (2 x 30MW) &bull; Section Electric, Instrument & Control (EIC)**

| Document Metric | Specification |
| :--- | :--- |
| **Document Version** | **2.2 (Annual Outage Rollover & Campaign Setup Release)** |
| **Author / Lead** | M. Farhan Tandia (EIC & IT Supervisor) & EIC Team PLTU MSW |
| **Last Updated** | **September 2, 2026** |
| **Status** | **Approved & Active in Production** |
| **Runtime Target** | Standalone Windows Executable (`server.exe`) & Python 3.8+ (Localhost / LAN) |

---

## 📜 Table of Contents
1. [Background & Problem Statement](#1-background--problem-statement)
2. [Product Objectives & Business Impact](#2-product-objectives--business-impact)
3. [System Architecture & Deployment](#3-system-architecture--deployment)
4. [Project Directory & File Structure](#4-project-directory--file-structure)
5. [Database Schema & Master Excel Specification](#5-database-schema--master-excel-specification)
6. [Functional Requirements (FR)](#6-functional-requirements)
7. [Non-Functional Requirements (NFR)](#7-non-functional-requirements)
8. [Design System & Modern Iconography](#8-design-system--modern-iconography)
9. [Complete Changelog & Evolution History](#9-complete-changelog--evolution-history)
10. [Field Operations & Standard Operating Procedures](#10-field-operations--standard-operating-procedures)

---

## 1. Background & Problem Statement

During periodic major and minor overhauls (*Major / Minor Outage*), the Section Electric, Instrument & Control (EIC) at PLTU MSW manages hundreds of high-precision field tasks. Historically, these records were fragmented across disparate spreadsheet files:

1. **`Progress_Outage_Unit_1_EIC.xlsx`** — Master Work Order tracker containing mechanical, electrical, and control subtasks (e.g., motor insulation resistance testing, contact resistance, relay calibrations, solo runs).
2. **`AMP-MSW-Progress_Actuator_Unit_1_2026.xlsx`** — Motorized actuator valve inspection tracker covering physical general inspections and function stroke tests across Boiler, Turbine, and Auxiliaries.
3. **`JAPA-MSW-Progress_Transmitter_&_Switch_Unit_2_2026.xls`** — Sensor calibration logs for Pressure Transmitters (PTX), Temperature Transmitters (TTX), and Pressure Switches (PSW) recording *As Found* vs. *As Left* data.
4. **`PIC_Scope_Master`** — Vendor vs. MSW scope delineation and personnel resource allocation.

### Operational Challenges:
- **Data Fragmentation & Inconsistency:** Manual consolidation across multiple spreadsheet formats introduced version conflicts, duplicate entries, and reporting lags.
- **Lack of Unified Real-Time Visibility:** Supervisors and plant management had to manually compile multiple workbooks to ascertain overall progress and identify bottlenecks.
- **Reporting Overhead:** Compiling daily progress reports, computing cumulative S-Curves, and formatting WhatsApp team briefings required significant manual effort.
- **Security & Authorization Risks:** Master scopes and personnel assignments lacked protection against accidental edits or unauthorized overwrites.

---

## 2. Product Objectives & Business Impact

The **Outage EIC Work Order Monitoring System** solves these challenges through a unified, high-performance platform:

- **Single Source of Truth:** Centralizes Work Orders, Actuator Valves, Field Instruments, Master Scopes, and Team PICs into structured master workbooks (`Template_Outage_EIC_Monitoring_unit 1.xlsx` and `Template_Outage_EIC_Monitoring_unit 2.xlsx`).
- **Two-Way Thread-Safe Synchronization:** Subtask toggles, finding logs, and status updates on the web UI commit instantly to Excel files with file-level locking (`threading.Lock`), ensuring data integrity.
- **Automated Multi-Format Reporting:** Instant 1-click generation of 4 standardized PDF reports, full `.xlsx` workbook exports, mathematical S-Curve trajectories, and WhatsApp briefing formats.
- **Master Security & Authorization:** Password-protected Master EIC & PIC scopes ensuring data governance.
- **Paperless Field Evidence:** Integrated field photo evidence upload, defect descriptions, and corrective action logging.
- **Enterprise-Grade UI/UX:** Full English localization, minimalist vector SVG iconography (Lucide style), dark/light themes, and responsive layout.

---

## 3. System Architecture & Deployment

The platform is designed with a **Zero-Dependency Portable Server Architecture** that runs natively in industrial power plant network environments:

```
+-----------------------------------------------------------------------------------+
|                                 CLIENT BROWSERS                                   |
|   (Desktop Workstations, Control Room PCs, Field Smartphones & Tablets via LAN)   |
+-----------------------------------------+-----------------------------------------+
                                          | HTTP / JSON API (Port 8000)
                                          v
+-----------------------------------------------------------------------------------+
|                            PYTHON BACKEND HTTP SERVER                             |
|              (Standalone `server.exe` or `python server.py`)                      |
|                                                                                   |
|  * Thread-Safe REST Endpoints:                                                    |
|    - `/api/data`               : Fetches full outage datasets & KPI metrics       |
|    - `/api/quick_toggle_subtask`: Real-time subtask check & cross-component sync  |
|    - `/api/update_wo`          : Work Order modifications & checklist persistence |
|    - `/api/update_actuator`    : General inspection & function test updates       |
|    - `/api/update_instrument`  : Calibration & verification milestones            |
|    - `/api/upload_finding_photo`: Multi-photo upload & defect logging             |
|    - `/api/export_excel`       : Live Excel workbook export (.xlsx)               |
|                                                                                   |
|  * Business Logic Engines:                                                        |
|    - Smart KKS & Tag Matcher   : Cross-component synchronization engine           |
|    - S-Curve Math Generator    : Normalized sigmoid trajectory & variance logic   |
|    - Excel Read/Write Engine   : Thread-locked openpyxl handler                   |
+-----------------------------------------+-----------------------------------------+
                                          | Safe File I/O
                                          v
+-----------------------------------------------------------------------------------+
|                               PERSISTENCE LAYER                                   |
|                                                                                   |
|  * Master Excel Workbooks:                                                        |
|    - `Template_Outage_EIC_Monitoring_unit 1.xlsx` (Unit 1 Data)                   |
|    - `Template_Outage_EIC_Monitoring_unit 2.xlsx` (Unit 2 Data)                   |
|  * Field Photos & Media:                                                          |
|    - `Finding/UNIT 1/<Equipment_ID>/`                                             |
|    - `Finding/UNIT 2/<Equipment_ID>/`                                             |
+-----------------------------------------------------------------------------------+
```

---

## 4. Project Directory & File Structure

```plaintext
d:\GitHub\msw_eic_om\
├── server.exe                                    # Standalone Windows executable (No Python runtime required)
├── server.py                                     # Core Python server, REST API, & Single Page Application
├── start_app.bat                                 # One-click startup batch script
├── server.spec                                   # PyInstaller compilation specification
├── Template_Outage_EIC_Monitoring_unit 1.xlsx    # Master database workbook for Unit 1
├── Template_Outage_EIC_Monitoring_unit 2.xlsx    # Master database workbook for Unit 2
├── Finding/                                      # Persistent storage for photo evidence
│   ├── UNIT 1/
│   │   └── <Equipment_ID>/foto_1.jpg, foto_2.jpg, ...
│   └── UNIT 2/
│       └── <Equipment_ID>/foto_1.jpg, foto_2.jpg, ...
├── screenshots/                                  # Visual documentation & system screenshots
├── README.md                                     # Main repository documentation & guide
└── PRD_Outage_EIC_WO_Monitoring.md               # Product Requirements Document (PRD v2.0)
```

---

## 5. Database Schema & Master Excel Specification

Each master workbook contains **8 standardized sheets**:

| Sheet Name | Purpose | Primary Data Columns |
| :--- | :--- | :--- |
| **`WorkOrder`** | Primary WO records | `No`, `No_WO`, `Unit`, `Job_Description`, `Area`, `Schedule_Date`, `Actual_Start_Date`, `Finish_Date`, `Status`, `PIC`, `N_Task`, `Progress_Percent`, `Scope`, `Remarks`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`WorkOrder_Checklist`** | Technical subtask checklists | `No_WO`, `Sub_Task_Description`, `Date`, `PIC_Task`, `Done_TRUE_FALSE`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`ActuatorValve`** | Motorized actuator valve matrix | `Equipment_ID`, `Area`, `Equipment_Description`, `KKS`, `Unit`, `PIC`, `Status`, `Progress_Percent`, `Finish_Date`, `General_Inspection_TRUE_FALSE`, `Function_Test_TRUE_FALSE`, `Remarks`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`Instrument_PressureTX`** | Pressure transmitter calibrations | `No`, `Area`, `Equipment`, `KKS`, `Unit`, `Range`, `Date`, `Finish_Date`, `Done_TRUE_FALSE`, `Remarks`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`Instrument_TemperatureTX`** | Temperature transmitter calibrations | `No`, `Area`, `Equipment`, `KKS`, `Unit`, `Range`, `Date`, `Finish_Date`, `Done_TRUE_FALSE`, `Remarks`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`Instrument_PressureSwitch`** | Pressure switch set point testing | `No`, `Area`, `Description`, `KKS`, `Unit`, `Sub_Area`, `Set_Point`, `Contact_Type_NO_NC`, `AsFound_Set`, `AsFound_Reset`, `AsLeft_Set`, `AsLeft_Reset`, `Status_OK_NotOK`, `Done_TRUE_FALSE`, `Date`, `Finish_Date`, `Remarks`, `Findings`, `Action_Taken`, `Photo_Count` |
| **`PIC_Scope_Master`** | Job scope & PIC mapping | `Category`, `Equipment_Scope_Name`, `Scope_Type_Vendor_MSW`, `Work_Scope_ME_SI_SE`, `Activity_Description`, `PIC`, `Unit` |
| **`Dashboard_Summary`** | Automated formula calculations & Outage Schedule | Formula-based summary (`COUNTIFS`) computing progress percentages & Section 5 storing official Outage Start Date and Outage Finish Date |

---

## 6. Functional Requirements

### 6.1. Real-Time Dashboard & Multi-Unit KPI Monitoring
- **FR-1.1 Unit Switcher:** Instant toggle between **UNIT 1** and **UNIT 2** with seamless state re-render.
- **FR-1.2 Grand Progress KPI:** Computes overall outage progress based on completed subtasks across all domains.
- **FR-1.3 Domain KPI Cards:** Separate visual progress metrics for Work Orders, Actuator Valves, Instruments, and Active Findings.
- **FR-1.4 Quick Filter Chips:** Instant 1-click filtering: *All Items*, *Active Findings / Photos*, *In Progress*, and *Completed*.
- **FR-1.5 Dynamic Pagination & Sizing:** Configurable pagination (**20 / page [Default]**, **40 / page**, **All**) with instant silent data refresh.

### 6.2. Work Order & Technical Subtask Management
- **FR-2.1 Collapsible Cards:** Accordion layout with collapsed default state for clear visual hierarchy.
- **FR-2.2 Interactive Checklists:** Instant check toggling with automatic completion timestamping (`DD/MM/YYYY`).
- **FR-2.3 Bulk Actions:** `Mark All Done` (sets 100% Finish) and `Reset` (clears checklist to 0% Sched-OK).
- **FR-2.4 Multi-Mode Subtask Addition:** Add subtasks via *Manual text input*, *Master Actuator picker*, or *Master Instrument picker*.
- **FR-2.5 Auto Finish Date:** Automatically sets completion date upon 100% checklist completion and clears it if unchecked.

### 6.3. Smart Cross-Component Two-Way Synchronization
- **FR-3.1 KKS Tag Normalization:** Intelligent matching engine reconciling prefix variations (`10` vs `20`) and technical synonyms (`DRAUGHT` &harr; `DRAFT`, `ID FAN` &harr; `INDUCED DRAUGHT FAN`).
- **FR-3.2 Bidirectional Propagations:** Checking a valve/instrument subtask in a Work Order automatically advances the corresponding component in `ActuatorValve` or `Instrument_*` sheet to **FINISH (100%)**, and vice-versa.

### 6.4. Master EIC Security & Authorization
- **FR-4.1 Default Protected State:** Master PIC personnel and Job Scope definitions are locked by default to prevent unauthorized modification.
- **FR-4.2 Clean Locked Mode:** Edit and delete action controls are completely hidden when locked.
- **FR-4.3 Password Authorization:** Modal unlock supporting master passwords with interactive SVG eye visibility toggle.
- **FR-4.4 Safe Lock Mode:** Allows one-click relocking after changes are completed.

### 6.5. Official Report Center & PDF Print Layouts
- **FR-5.1 Dedicated Report Center:** Modal interface providing 4 standardized report templates:
  1. *Report 1: Daily Progress & Findings Log* (with date range filtering).
  2. *Report 2: Full Work Orders & Detailed Subtasks*.
  3. *Report 3: Actuator Valves Inspection Matrix*.
  4. *Report 4: Field Instruments Calibration (PTX, TTX, PSW)*.
- **FR-5.2 Direct Excel Export:** Endpoint `/api/export_excel?unit=X` generates live, full-workbook `.xlsx` downloads.
- **FR-5.3 Print-Optimized Styling:** CSS `@media print` rules ensure clean page breaks, full-width tables, and high-contrast typography.

### 6.6. S-Curve Trajectory & Progress Analytics
- **FR-6.1 Mathematical S-Curve Model:** Calculates planned target trajectory using a normalized sigmoid function ($k=7.0$).
- **FR-6.2 Variance Analytics:** Displays real-time progress variance (`+X% Ahead` or `-X% Behind`).
- **FR-6.3 Two-Way Excel Outage Window Persistence:** Official Outage Start Date and Outage Finish Date stored directly in `Dashboard_Summary` Section 5 of both master Excel workbooks. Date edits in the web UI commit directly to Excel via `/api/update_outage_dates`, keeping the planned S-Curve trajectory stable and consistent across all devices and client sessions.
- **FR-6.4 Outage Banner Period Badge:** Interactive badge displayed in the Outage Banner showing active date window with 1-click modal access.

### 6.7. WhatsApp Coordination Briefing Generator
- **FR-7.1 Structured Message Formatter:** Automatically compiles Grand Progress, Work Orders, Actuators, Instruments, tasks completed today, and open defect findings.
- **FR-7.2 1-Click Clipboard Copy:** Instant copy button for sharing to team coordination groups.

### 6.8. Field Evidence & Photo Management
- **FR-8.1 Defect & Action Logging:** Dual textarea inputs for abnormal findings and recommended corrective action plans.
- **FR-8.2 Multi-Photo Evidence:** Drag-and-drop file uploader and direct camera capture (`accept="image/*" capture="environment"`).
- **FR-8.3 Unit & Component-Named Storage Architecture:** Photos and field findings are cleanly separated per unit into `Finding/UNIT 1/` and `Finding/UNIT 2/`. Folder names use descriptive human-readable component names (e.g. `ACTUATOR TURBINE CONTROL`, `PRESSURE TRANSMITTER - MAIN STEAM`, `OM EIC Y1 UNIT COMMON COOLING TOWER FAN 1`) rather than raw KKS or equipment IDs, synchronized with Excel `Photo_Count`.
- **FR-8.4 Fullscreen Lightbox:** Click-to-enlarge modal for field photo inspections.

### 6.9. Annual Outage Rollover & Campaign Setup
- **FR-9.1 Password-Gated Access Control:** Access to the Outage Rollover Wizard is strictly restricted behind Master Authorization Password validation at both frontend UI and backend API layers.
- **FR-9.2 Work Order Mapping Engine:** Endpoint `/api/export_wo_mapping?unit=X` dynamically generates Excel templates containing current active WOs for easy assignment of new WO numbers from CMMS (SAP / Maximo).
- **FR-9.3 Flexible Remapping Modes:** Supports both *Upload Filled Excel Mapping* and *Quick Prefix Replace* (e.g. `WO-100826-` &rarr; `WO-150827-`).
- **FR-9.4 Cascading Data Integrity:** Remaps Work Orders and replicates new WO numbers across all 760+ checklist subtasks without broken links.
- **FR-9.5 Automated Archiving & Zero Data Loss:** Automatically creates timestamped archives of existing workbooks and photo documentation in `Archive/` before resetting progress to 0% (`SCHED-OK` status, reset checkboxes, updated S-Curve baseline dates).

---

## 7. Non-Functional Requirements

- **Performance & Latency:** Excel read/write commits execute in $< 500\text{ ms}$.
- **Data Concurrency:** File locking prevents spreadsheet corruption during simultaneous user access.
- **Portability & Deployment:** Standalone binary (`server.exe`) runs without external dependencies.
- **Cross-Device Usability:** Responsive viewport design supporting desktop workstations, laptops, tablets, and smartphones ($360\text{px} - 430\text{px}$).

---

## 8. Design System & Modern Iconography

The interface adheres to modern enterprise SaaS standards:

- **Vector SVG Icons:** Centralized stroke-based SVG icon library (`1.8px` stroke width, `currentColor` theme inheritance) completely replacing legacy OS emojis.
- **Industrial Color Palette:**
  - **Primary:** Electric Indigo (`#6366f1` / `#4f46e5`)
  - **Finish Status:** Emerald Green (`#10b981`)
  - **In Progress:** Sky Blue (`#38bdf8`)
  - **Scheduled / Pending:** Slate Gray (`#94a3b8`)
  - **Alerts / Findings:** Rose Crimson (`#f43f5e`)
- **Dual Themes:** Clean Dark Mode (default for control rooms) and High-Contrast Light Mode with persistent `localStorage` state.
- **Typography:** `Inter` for interface elements and `JetBrains Mono` for KKS codes, dates, and numerical metrics.

---

## 9. Complete Changelog & Evolution History

### 🔹 Version 2.2 (September 2, 2026) — *Annual Outage Rollover & Progress Reset Wizard*
- **Password-Protected Outage Rollover Wizard:** Added a dedicated modal wizard accessible via the header `Rollover` button, guarded by Master EIC Password authentication.
- **Dual Remapping Architecture:** Added support for both uploaded Excel mapping files (`WO_Mapping_Unit_X.xlsx`) and 1-click Quick Prefix replacement.
- **Area & Job Description Remapping:** Synchronizes modifications made to `Area` and `Job_Description` in mapping files directly into master workbooks.
- **Unit & Component-Named Finding Organization:** Restructured finding photo directories to `Finding/UNIT 1/` and `Finding/UNIT 2/`, with folders named by descriptive component names (e.g. `ACTUATOR TURBINE CONTROL`) instead of raw KKS or equipment codes.
- **Automated Cascading WO Update & Archiving:** Automatically backs up master Excel workbooks and photos to `Archive/`, cascades new WO numbers into `WorkOrder` and `WorkOrder_Checklist`, and resets all progress, checklists, actuators, instruments, and findings to 0% (`SCHED-OK`).
- **Endpoints `/api/export_wo_mapping` & `/api/rollover_outage`:** Thread-safe endpoints for template generation and multi-unit rollover execution.

### 🔹 Version 2.1 (September 2, 2026) — *Excel Outage Schedule Persistence & Stable S-Curve*
- **Two-Way Excel Outage Schedule Synchronization:** Added Section 5 (`OUTAGE SCHEDULE & S-CURVE PERIOD`) in `Dashboard_Summary` sheet across Unit 1 and Unit 2 master workbooks storing official Outage Start Date and Outage Finish Date.
- **API Endpoint `/api/update_outage_dates`:** Thread-safe backend endpoint allowing users to modify outage dates from the web UI and immediately write changes into the Excel file.
- **Outage Banner Period Badge:** Added interactive Outage Period badge in the dashboard banner with 1-click shortcut to open and adjust S-Curve dates.
- **Eliminated S-Curve Date Drift:** Resolved volatile date shifting caused by dynamic task date discovery by anchoring calculations to official Excel outage dates.

### 🔹 Version 2.0 (September 1, 2026) — *Enterprise Iconography & Full Localization*
- **Vector SVG Icon System:** Replaced 50+ OS emoji characters with crisp, stroke-based inline SVG icons (Lucide style) across all headers, tabs, cards, tables, buttons, and modals.
- **100% Full English Localization:** Complete translation of backend API responses, frontend HTML/JS, report templates, and master Excel workbooks (`Template_Outage_EIC_Monitoring_unit 1.xlsx` and `unit 2.xlsx`) across all 8 sheets.
- **Master EIC Security & Clean Locked Mode:** Added Master Authorization password protection for Scope and Master PIC management. Action buttons are hidden when locked, with an SVG eye toggle for password visibility.
- **In-Place Silent Data Refresh:** Implemented a silent refresh button beside the pagination indicator, updating data without page reloads or loss of scroll position.
- **Default Status Scheduled:** Standardized default status for newly created items to `SCHED-OK` (0% progress).
- **Responsive Mobile Styling:** Enhanced adaptive styling for smartphones and tablets.

### 🔹 Version 1.5 (August 30, 2026)
- **Dynamic Pagination:** Added page size options (**20 / page**, **40 / page**, **All**).
- **Accordion Initialization:** Initialized all cards in collapsed state for improved visual hierarchy.
- **Interactive Header Logo:** Clicking the header logo triggers an instant dashboard reload.

### 🔹 Version 1.4 (August 30, 2026)
- **Cross-Component Smart Sync:** Implemented intelligent KKS matching connecting Work Order checklists with Actuator and Instrument sheets in real-time.
- **Instant Background Toggle:** Subtask checkboxes save directly to `/api/quick_toggle_subtask` without full page refresh.

### 🔹 Version 1.0 - 1.3 (August 29–30, 2026)
- **Centralized Reports Center:** Unified 4 official report formats, S-Curve trends, WhatsApp generator, and Excel download into a dedicated modal.
- **Unit Switcher Relocation:** Relocated Unit 1 / Unit 2 switchers into the Outage Banner.
- **Sticky Summary Bar & Back-to-Top:** Added floating progress bar and scroll-to-top button.

### 🔹 Version 0.1 - 0.9 (August 29, 2026)
- **Initial Architecture:** Built pure Python HTTP server, openpyxl Excel handler, REST endpoints, dual themes, and initial PRD specifications.

---

## 10. Field Operations & Standard Operating Procedures

1. **Starting the System:**
   - Double-click `server.exe` (or run `start_app.bat`).
   - The application automatically launches on `http://localhost:8000`.
2. **Network Access for Field Technicians:**
   - Connect client smartphones/tablets to the power plant Wi-Fi / LAN network.
   - Navigate to `http://<SERVER_IP>:8000` (e.g., `http://192.168.1.104:8000`).
3. **Updating Work Orders & Subtasks:**
   - Select the target Work Order card to expand the checklist.
   - Check completed subtasks; completion dates and linked actuator/instrument records sync automatically.
4. **Logging Field Findings & Photos:**
   - Click the camera button on any item.
   - Attach photos, enter defect details and corrective action plans, then save.
5. **Exporting Reports & WhatsApp Briefings:**
   - Click **Reports** in the header.
   - Select the desired report format, adjust date filters, and click **Print / Export PDF** or **WhatsApp Summary**.
6. **Annual Outage Rollover (Transitioning to Next Year):**
   - Click **Rollover** in the header controls.
   - Enter the Master EIC Password to unlock the wizard.
   - Select the target Unit (Unit 1, Unit 2, or Both).
   - Download the WO Mapping Template to review or fill new WO numbers, or use Quick Prefix Replace if only date codes change.
   - Set the New Outage Start and Finish baseline dates.
   - Click **Eksekusi Outage Rollover & Reset Progress**. The system automatically archives existing data, updates WO numbers, resets all progress to 0%, and refreshes the live dashboard.

---

<div align="center">

**PLTU MSW &bull; Section Electric, Instrument & Control**  
*Excellence in Operation & Maintenance Execution*

</div>
