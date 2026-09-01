<div align="center">

# Outage EIC Work Order Monitoring System
### Real-Time Overhaul Monitoring & Field Management Platform &bull; Unit 1 & Unit 2 (2 x 30MW)
**PLTU MSW &bull; Section Electric, Instrument & Control (EIC)**

[![Status](https://img.shields.io/badge/Status-Production%20Ready-10b981?style=for-the-badge)](https://github.com/)
[![Version](https://img.shields.io/badge/Version-2.0%20Enterprise-6366f1?style=for-the-badge)](https://github.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Web%20%7C%20Mobile-2563eb?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/)
[![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Excel Sync](https://img.shields.io/badge/Excel%20Sync-openpyxl%20Thread--Safe-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://github.com/)

<br />

<img src="./screenshots/dashboard_dark.png" alt="PLTU MSW Outage EIC Dashboard" width="100%" style="border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);" />

</div>

---

## 📌 Table of Contents
- [1. Executive Summary & Background](#1-executive-summary--background)
- [2. Key System Capabilities](#2-key-system-capabilities)
- [3. Feature Deep Dive & Architecture](#3-feature-deep-dive--architecture)
  - [3.1. Real-Time Multi-Unit KPI Dashboard](#31-real-time-multi-unit-kpi-dashboard)
  - [3.2. Work Orders & Subtask Checklists](#32-work-orders--subtask-checklists)
  - [3.3. Actuator Valve Monitoring](#33-actuator-valve-monitoring)
  - [3.4. Field Instrument Calibration (PTX, TTX, PSW)](#34-field-instrument-calibration-ptx-ttx-psw)
  - [3.5. Master EIC Scope Security & Password Authorization](#35-master-eic-scope-security--password-authorization)
  - [3.6. Official Report Center & Live Excel Export](#36-official-report-center--live-excel-export)
  - [3.7. S-Curve Progress Analytics & Variance](#37-s-curve-progress-analytics--variance)
  - [3.8. Structured WhatsApp Coordination Briefing](#38-structured-whatsapp-coordination-briefing)
  - [3.9. Field Evidence & Photo Management](#39-field-evidence--photo-management)
  - [3.10. Modern Vector SVG Icon System & Dual Themes](#310-modern-vector-svg-icon-system--dual-themes)
- [4. Repository & File Structure](#4-repository--file-structure)
- [5. Getting Started (Running the Application)](#5-getting-started-running-the-application)
- [6. Multi-Device Network Access (LAN / Mobile)](#6-multi-device-network-access-lan--mobile)
- [7. Excel Database Protocol & Concurrency](#7-excel-database-protocol--concurrency)
- [8. Troubleshooting & FAQ](#8-troubleshooting--faq)

---

## 1. Executive Summary & Background

The **Outage EIC Work Order Monitoring System** is an industrial-grade periodic maintenance (*Overhaul / Outage*) monitoring platform engineered specifically for the **Section Electric, Instrument & Control (EIC) at PLTU MSW (2 x 30MW)**.

During outages, maintenance engineers execute hundreds of critical field tasks across Work Orders, Motorized Actuator Valves, Transmitters, and Pressure Switches. This platform replaces fragmented spreadsheets with a **centralized single-source-of-truth dashboard** that synchronizes bidirectionally in real time with standardized Excel workbooks (`Template_Outage_EIC_Monitoring_unit 1.xlsx` and `Template_Outage_EIC_Monitoring_unit 2.xlsx`).

---

## 2. Key System Capabilities

- **Real-Time Visibility:** Live overall progress computation (*Grand Total, Work Orders, Actuators, and Instruments*) across Unit 1 and Unit 2.
- **Two-Way Thread-Safe Sync:** Web checklist toggles commit instantly to Excel master workbooks with file locking (`threading.Lock`), ensuring zero data corruption.
- **Automated Operational Reporting:** 1-click generation of 4 official PDF print-ready reports, full `.xlsx` workbook exports, mathematical S-Curve trajectories, and formatted WhatsApp team briefings.
- **Master Security & Governance:** Password-protected Master EIC PIC personnel and Job Scope definitions to prevent accidental overwrites.
- **Paperless Field Documentation:** Multi-photo attachment upload, defect finding logging, and corrective action tracking.
- **Zero-Dependency Portability:** Standalone binary (`server.exe`) running natively on Windows workstations and accessible via local LAN by smartphones and tablets.

---

## 3. Feature Deep Dive & Architecture

### 3.1. Real-Time Multi-Unit KPI Dashboard
The dashboard delivers instant high-level situational awareness:
- **Outage Progress Banner:** Shows overall completion percentage (*Grand Total*), total task counts, and instant toggle between **UNIT 1** and **UNIT 2**.
- **Domain KPI Summary Cards:**
  - **Work Orders:** Total WOs, completed subtasks, and in-progress jobs.
  - **Actuator Valves:** Readiness status for General Inspection and Function Test.
  - **Instruments:** Verification status for Pressure TX, Temperature TX, and Pressure Switches.
  - **Active Findings:** Counts open field defects requiring resolution.
- **Sticky Summary Bar:** Automatically appears upon scrolling past 180px, displaying key metrics, a quick report trigger, and a scroll-to-top button.
- **Toolbar & Filter Pills:** Clean search by WO, KKS, Equipment, Area, Findings, or PIC, paired with quick status filters (*All Items*, *Active Findings / Photos*, *In Progress*, *Completed*).
- **In-Place Silent Refresh:** Re-fetches the latest Excel state silently without page reloads or losing scroll position.

<div align="center">
  <img src="./screenshots/dashboard_dark.png" alt="Dashboard Dark Mode" width="100%" />
</div>

---

### 3.2. Work Orders & Subtask Checklists
Each Work Order card features an accordion layout:
- **Interactive Checklists:** Subtasks can be checked individually, automatically recording today's date (`DD/MM/YYYY`) and updating progress percentage.
- **Bulk Actions:**
  - **`Mark All Done`**: Completes all subtasks simultaneously and sets WO status to **100% FINISH**.
  - **`Reset`**: Resets all checkboxes back to **0% SCHED-OK**.
- **Multi-Mode Subtask Addition:**
  - *Manual Input:* Type custom subtask descriptions.
  - *Select Actuator:* Pick from registered master actuators (with automatic KKS and area linking).
  - *Select Instrument:* Pick from master transmitters or pressure switches.
- **Auto Finish Date:** Automatically sets completion date when all subtasks reach 100%, and clears it if any item is unchecked.

<div align="center">
  <img src="./screenshots/work_order_expanded.png" alt="Work Order Expanded View" width="100%" />
</div>

---

### 3.3. Actuator Valve Monitoring
Dedicated matrix tracking all motorized actuator valves across Boiler, Turbine, and Auxiliary systems:
- **Two-Stage Milestone Tracking:** Independent verification for **General Inspection** (50%) and **Function Test** (50%).
- **Cards & Table View Modes:** Switch between card view and compact table view for batch verification.

<div align="center">
  <img src="./screenshots/actuator_valves_tab.png" alt="Actuator Valves Tab" width="100%" />
</div>

---

### 3.4. Field Instrument Calibration (PTX, TTX, PSW)
Comprehensive calibration tracking segmented into 3 sub-tabs:
- **Pressure Transmitters (PTX):** Tag KKS, Equipment, Range, Calibration status, and Verification (Done) milestone.
- **Temperature Transmitters (TTX):** Temperature sensor calibration logs.
- **Pressure Switches (PSW):** Complete set point data (`HIGH`/`LOW`), Contact Type (`NO`/`NC`), *As Found* set/reset, *As Left* set/reset, and Pass/Fail result.

<div align="center">
  <img src="./screenshots/instruments_tab.png" alt="Instruments Tab" width="100%" />
</div>

---

### 3.5. Master EIC Scope Security & Password Authorization
Protects master data integrity:
- **Locked Mode by Default:** Master PIC personnel and Job Scope definitions are locked to prevent unauthorized modification. Edit and delete buttons are cleanly hidden.
- **Password Authorization Modal:** Authorizes supervisors using master credentials (`eic123`, `admin123`, `msweic`, etc.) with an interactive SVG Eye toggle.
- **Safe Relock:** Easily relock access with 1 click after modifications are saved.

---

### 3.6. Official Report Center & Live Excel Export
Accessible via the **`Reports`** button in the header, providing 4 standardized templates:
1. **Report 1: Daily Progress & Findings:** Summary of subtasks completed within a specified date range along with active findings.
2. **Report 2: Full WO & Sub-Tasks:** Complete breakdown of all Work Orders, subtasks, and assigned PICs.
3. **Report 3: Actuator Valves:** Inspection and function test matrix for all valve actuators.
4. **Report 4: Field Instruments:** Verification summary for PTX, TTX, and PSW instruments.
5. **Download Excel (.xlsx):** Live full-workbook download synchronized with the database.
6. **Print / Export PDF:** Print-optimized layout with clean page breaks and high-contrast typography.

<div align="center">
  <img src="./screenshots/report_modal.png" alt="Report Modal" width="100%" />
</div>

---

### 3.7. S-Curve Progress Analytics & Variance
Mathematical S-Curve visualization engine:
- **Sigmoid Trajectory:** Models planned baseline trajectory using a normalized sigmoid function ($k=7.0$).
- **Variance KPI:** Computes actual progress against the planned baseline (`+X% Ahead` or `-X% Behind`).
- **Configurable Outage Window:** Date pickers saved in `localStorage` for dynamic schedule adjustments.

<div align="center">
  <img src="./screenshots/s_curve_graph.png" alt="S-Curve Graph Modal" width="100%" />
</div>

---

### 3.8. Structured WhatsApp Coordination Briefing
Generates structured summary text for WhatsApp team coordination:
- Summarizes Grand Total progress, WO completion ratios, Actuator readiness, Instrument verification, tasks completed today, and open defect findings.
- Includes a **1-Click "Copy to Clipboard"** button.

<div align="center">
  <img src="./screenshots/whatsapp_format.png" alt="WhatsApp Generator Modal" width="100%" />
</div>

---

### 3.9. Field Evidence & Photo Management
- Multi-photo attachment supporting JPEG, PNG, and WebP via drag-and-drop or camera capture.
- Photos stored in `Finding/UNIT X/<Item ID>/` and synchronized with Excel `Photo_Count`.
- Fullscreen lightbox preview with delete capability.

---

### 3.10. Modern Vector SVG Icon System & Dual Themes
- **Clean Vector Iconography:** Uses lightweight, stroke-based SVG icons (Lucide style) with `currentColor` inheritance, eliminating legacy OS emoji clutter.
- **Theme Switcher:** Clean Sun / Moon SVG toggle supporting Dark Mode and Light Mode.

<div align="center">
  <img src="./screenshots/light_mode.png" alt="Light Mode Dashboard" width="100%" />
</div>

---

## 4. Repository & File Structure

```plaintext
d:\GitHub\msw_eic_om\
├── server.exe                                    # Standalone Windows executable (No Python required)
├── server.py                                     # Python backend server, REST API, & Single Page App
├── start_app.bat                                 # Quick launcher script
├── server.spec                                   # PyInstaller build specification
├── Template_Outage_EIC_Monitoring_unit 1.xlsx    # Master database workbook for Unit 1
├── Template_Outage_EIC_Monitoring_unit 2.xlsx    # Master database workbook for Unit 2
├── Finding/                                      # Storage directory for field evidence photos
│   ├── UNIT 1/
│   │   └── <Equipment_ID>/foto_1.jpg, foto_2.jpg, ...
│   └── UNIT 2/
│       └── <Equipment_ID>/foto_1.jpg, foto_2.jpg, ...
├── screenshots/                                  # Documentation screenshots
├── README.md                                     # Main repository guide
└── PRD_Outage_EIC_WO_Monitoring.md               # Product Requirements Document (PRD v2.0)
```

---

## 5. Getting Started (Running the Application)

### Option 1: Standalone Executable (Recommended for Field Users)
Double-click `server.exe`. The application starts on port `8000` and opens your default browser at:
```plaintext
http://localhost:8000
```

### Option 2: Quick Launcher Batch Script
Double-click `start_app.bat`.

### Option 3: Python Source Code (Development)
1. Install Python dependencies:
   ```bash
   pip install openpyxl pillow
   ```
2. Start the server:
   ```bash
   python server.py
   ```

### Option 4: Recompile Executable (.exe)
If modifications are made to `server.py`, recompile using PyInstaller:
```powershell
pyinstaller server.spec --distpath . -y
```

---

## 6. Multi-Device Network Access (LAN / Mobile)

To access the dashboard from smartphones, tablets, or other workstations on the plant network:
1. Ensure the device is connected to the same Wi-Fi / LAN network as the host server machine.
2. Find the host server machine's IP address (e.g., `192.168.1.104`).
3. Navigate to:
   ```plaintext
   http://192.168.1.104:8000
   ```

---

## 7. Excel Database Protocol & Concurrency

The system uses safe file-locking protocols (`threading.Lock`) for concurrent read/write operations:

| Sheet Name | Data Description | Sync Behavior |
| :--- | :--- | :--- |
| **`WorkOrder`** | WO No, Description, Area, PIC, Status, % Progress, Finish Date, Remarks | Updated instantly when subtasks or general data change. |
| **`WorkOrder_Checklist`** | Technical subtask checklists, Dates, Task PICs, Completion (TRUE/FALSE) | Real-time bidirectional sync with Actuators & Instruments. |
| **`ActuatorValve`** | Equipment ID, KKS, General Inspection, Function Test, Status | Synchronized when actuator subtasks in WO are toggled. |
| **`Instrument_*`** | KKS Tag, Equipment, Range/Set Point, Calibration & Verification | Synchronized when instrument subtasks in WO are toggled. |
| **`PIC_Scope_Master`** | Job scope delineations, Scope types (Vendor/MSW), Assigned PICs | Synchronized when Master Scope rows are edited or added. |
| **`Dashboard_Summary`** | Automatic formula calculations | Formula-based summary (`COUNTIFS`) computing progress percentages. |

---

## 8. Troubleshooting & FAQ

<details>
<summary><b>❓ Error: "Excel file is locked / Permission denied" when saving</b></summary>

Close the Excel spreadsheet `Template_Outage_EIC_Monitoring_unit X.xlsx` if it is currently open in Microsoft Excel desktop so the server process can write updates safely.
</details>

<details>
<summary><b>❓ How do I refresh the data without losing my current view?</b></summary>

Click the **Refresh** button next to the pagination counter (`Showing X-Y of Z items`) to reload data from Excel in place.
</details>

<details>
<summary><b>❓ What is the Master Authorization password?</b></summary>

Default master passwords include `eic123`, `admin123`, or `msweic`.
</details>

---

<div align="center">

**PLTU MSW &bull; Section Electric, Instrument & Control (EIC)**  
*Continuous Improvement &bull; Excellence in Plant Reliability & Safety*

</div>
