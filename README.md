<div align="center">

# ⚡ PLTU MSW EIC Outage Monitoring System
### Real-Time Monitoring & Work Order Management Platform &bull; Unit 1 & Unit 2 (2 x 30MW)

[![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-10b981?style=for-the-badge)](https://github.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Web-2563eb?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/)
[![Section](https://img.shields.io/badge/Section-Electric%20Instrument%20%26%20Control-f59e0b?style=for-the-badge)](https://github.com/)
[![Excel Sync](https://img.shields.io/badge/Data%20Sync-Excel%20openpyxl-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://github.com/)

<br />

<img src="./screenshots/dashboard_dark.png" alt="PLTU MSW Outage EIC Dashboard" width="100%" style="border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);" />

</div>

---

## 📌 Table of Contents
- [1. Introduction & Background](#-1-introduction--background)
- [2. System Objectives](#-2-system-objectives)
- [3. Key Features & Visual Documentation](#-3-key-features--visual-documentation)
  - [3.1. Real-Time Dashboard & KPI Progress](#31-real-time-dashboard--kpi-progress)
  - [3.2. Work Order Management & Batch Actions](#32-work-order-management--batch-actions)
  - [3.3. Actuator Valve Monitoring](#33-actuator-valve-monitoring)
  - [3.4. Instrument Verification & Calibration](#34-instrument-verification--calibration)
  - [3.5. Official Report Center & Export](#35-official-report-center--export)
  - [3.6. S-Curve Analysis & Daily Trends](#36-s-curve-analysis--daily-trends)
  - [3.7. WhatsApp Summary Generator](#37-whatsapp-summary-generator)
  - [3.8. Dual Theme (Dark & Light Mode)](#38-dual-theme-dark--light-mode)
- [4. Directory Structure & Architecture](#-4-directory-structure--architecture)
- [5. Getting Started (Running the Application)](#-5-getting-started-running-the-application)
- [6. Excel Data Synchronization Protocol](#-6-excel-data-synchronization-protocol)
- [7. Troubleshooting & FAQ](#-7-troubleshooting--faq)

---

## 📖 1. Introduction & Background

**PLTU MSW EIC Outage Monitoring System** is a local web-based periodic maintenance (*Overhaul / Outage*) monitoring platform specifically engineered for the **Electric, Instrument & Control (EIC) Section at PLTU MSW (2 x 30MW)**.

The system facilitates maintenance teams in monitoring, updating, and documenting the progress of hundreds of field work items (Work Orders, Motorized Valve Actuators, and Sensor Transmitters & Switches) in real-time. It features bidirectional, real-time synchronization (*two-way sync*) with Microsoft Excel master spreadsheet databases (`Template_Outage_EIC_Monitoring_unit 1.xlsx` and `Template_Outage_EIC_Monitoring_unit 2.xlsx`).

---

## 🎯 2. System Objectives

1. **⚡ Real-Time Visibility:** Provides an accurate overview of daily progress percentages (*Grand Total, WO, Valves, & Instruments*) without requiring manual recap.
2. **🔄 Automatic 2-Way Sync:** Web checklist updates are instantly committed to Excel, and master Excel changes are dynamically read by the system.
3. **📷 Digital Field Evidence (*Paperless*):** Supports multi-photo uploads for findings, abnormality logging (*abnormal findings*), and follow-up action planning.
4. **📊 Industry Standard Reporting:** Offers 4 official PDF print-ready report formats, instant `.xlsx` spreadsheet exports, and structured WhatsApp daily briefing generators.

---

## 🚀 3. Key Features & Visual Documentation

### 3.1. Real-Time Dashboard & KPI Progress
Presents the outage achievement matrix with automatic KPI summary cards.

<div align="center">
  <img src="./screenshots/dashboard_dark.png" alt="Real-Time Dashboard Dark Mode" width="100%" />
</div>

- **Progress Banner:** Displays total achievement percentage (*Huge Metric*), completed task ratio, and unit selector buttons (**UNIT 1** / **UNIT 2**).
- **Vertical KPI Metrics:**
  - 📋 **Work Order:** Total completed WOs and completed sub-task ratio.
  - ⚙️ **Actuator Valves:** Readiness status for General Inspection & Function Test.
  - 🎛️ **Instruments:** Verification status for Pressure TX, Temperature TX, & Pressure Switch.
  - 🚨 **Field Findings:** Counts active anomalies that require resolution.

---

### 3.2. Work Order Management & Batch Actions
Each WO card includes work descriptions, plant areas, assigned PICs, and interactive sub-task checklists.

<div align="center">
  <img src="./screenshots/work_order_expanded.png" alt="Work Order Expanded View" width="100%" />
</div>

- **⚡ Batch / Bulk Actions:**
  - **`✓ Complete All`**: Checks all sub-tasks simultaneously, populates today's date, and advances the WO status directly to **100% FINISH**.
  - **`↺ Reset`**: Clears all checklist checkboxes for re-evaluation.
- **➕ Flexible Sub-Task Addition:** Supports entering descriptions *Manually*, selecting from *Master Actuators*, or selecting from *Master Instruments*.
- **📅 Auto-Fill Completion Date:** Finish date is automatically filled when all checklist items are completed, and automatically cleared if any item remains unchecked.

---

### 3.3. Actuator Valve Monitoring
Tracks all motorized actuator valves across Boiler, Turbine, and Auxiliary areas.

<div align="center">
  <img src="./screenshots/actuator_valves_tab.png" alt="Actuator Valves Tab" width="100%" />
</div>

- **Independent Inspection Checklists:** Separate verification for *General Inspection* (Physical/Mechanical/Electrical) and *Function Test* (Open/Close Stroke & DCS Feedback).
- **Comparison Matrix Mode:** Enables side-by-side comparison of actuator readiness between Unit 1 and Unit 2.

---

### 3.4. Instrument Verification & Calibration
Power plant instrumentation monitoring grouped into dedicated sub-tabs:
- **Pressure Transmitter (PTX)**
- **Temperature Transmitter (TTX)**
- **Pressure Switch (PSW)**

<div align="center">
  <img src="./screenshots/instruments_tab.png" alt="Instruments Tab" width="100%" />
</div>

---

### 3.5. Official Report Center & Export
Centralized access via the **`📑 Report`** button in the header, offering 4 standardized report formats:

<div align="center">
  <img src="./screenshots/report_modal.png" alt="Report Modal" width="100%" />
</div>

1. **Report 1: Daily Progress & Findings:** Summary of sub-tasks completed within a specific date range along with active findings.
2. **Report 2: Full WO & Sub-Tasks:** Comprehensive breakdown of all Work Orders, sub-tasks, and assigned PICs.
3. **Report 3: Actuator Valves:** Status matrix of all valve actuator inspections.
4. **Report 4: Instruments:** Verification summary for PTX, TTX, and PSW instruments.
5. **📥 Download Excel (.xlsx):** Download the latest live Excel database directly from the browser.
6. **🖨️ Print / Save as PDF:** Standardized operational document print layout.

---

### 3.6. S-Curve Analysis & Daily Trends
Interactive S-curve visualization chart to analyze actual daily progress velocity against planned outage completion targets.

<div align="center">
  <img src="./screenshots/s_curve_graph.png" alt="S-Curve Graph Modal" width="100%" />
</div>

---

### 3.7. WhatsApp Summary Generator
Automatic daily summary message generator structured cleanly for sharing directly to WhatsApp coordination groups.

<div align="center">
  <img src="./screenshots/whatsapp_format.png" alt="WhatsApp Generator Modal" width="100%" />
</div>

- Includes a 1-click **"📋 Copy to Clipboard"** button.

---

### 3.8. Dual Theme (Dark & Light Mode)
The interface supports dark and light mode switching via the icon toggle in the top-right header corner.

<div align="center">
  <img src="./screenshots/light_mode.png" alt="Light Mode Dashboard" width="100%" />
</div>

---

## 📁 4. Directory Structure & Architecture

```plaintext
d:\msw\msw_eic_om\
├── server.py                                    # Backend HTTP Server, JSON API, & Embedded UI Template
├── server.exe                                   # Standalone Binary Application (No Python installation required)
├── server.spec                                  # PyInstaller Build Configuration
├── start_app.bat                                # Quick Launcher Script
├── .gitignore                                   # Git Ignore Configuration
├── README.md                                    # Main Repository Documentation
├── Template_Outage_EIC_Monitoring_unit 1.xlsx   # Master Database Unit 1
├── Template_Outage_EIC_Monitoring_unit 2.xlsx   # Master Database Unit 2
├── Finding/                                     # Field Evidence & Anomaly Photo Storage
│   ├── UNIT 1/
│   └── UNIT 2/
└── screenshots/                                 # System Documentation Screenshots
    ├── dashboard_dark.png
    ├── work_order_expanded.png
    ├── actuator_valves_tab.png
    ├── instruments_tab.png
    ├── report_modal.png
    ├── s_curve_graph.png
    ├── whatsapp_format.png
    └── light_mode.png
```

---

## 🛠️ 5. Getting Started (Running the Application)

<details open>
<summary><b>🔹 Option 1: Run Standalone Binary (Recommended for Users / Technicians)</b></summary>

Simply double-click:
```plaintext
server.exe
```
The application will launch the local server on port 8000 and automatically open the dashboard in your default browser:
```plaintext
http://localhost:8000
```
</details>

<details>
<summary><b>🔹 Option 2: Run via Batch File Launcher</b></summary>

Double-click:
```plaintext
start_app.bat
```
</details>

<details>
<summary><b>🔹 Option 3: Run from Python Source Code (Development)</b></summary>

1. Ensure Python dependencies are installed:
   ```bash
   pip install openpyxl pillow
   ```
2. Start the server:
   ```bash
   python server.py
   ```
</details>

<details>
<summary><b>🔹 Option 4: Recompile to Executable (.exe)</b></summary>

If you make modifications to `server.py`, recompile using:
```powershell
pyinstaller server.spec --distpath . -y
```
</details>

---

## 🔄 6. Excel Data Synchronization Protocol

The system utilizes a thread-safe file locking mechanism to guarantee data consistency and integrity:

| Worksheet (Sheet) | Data Stored | Synchronization Mechanism |
| :--- | :--- | :--- |
| **`WorkOrder`** | WO No, Description, Area, PIC, Status, % Progress, Finish Date, Remarks | Automatically updated when sub-task checklists are toggled. |
| **`WorkOrder_Checklist`** | WO Sub-tasks, Task PIC, Completion Status (TRUE/FALSE), Date | Real-time bidirectional sync with Actuators & Instruments. |
| **`ActuatorValve`** | Equipment ID, Area, KKS, General Inspection, Function Test, Status | Automatically synced when actuator sub-tasks in WO are checked. |
| **`Instrument_*`** | No, KKS Tag, Equipment, Range, Calibration & Verification Status | Automatically synced when instrument sub-tasks in WO are checked. |

---

## 💡 7. Troubleshooting & FAQ

<details>
<summary><b>❓ How do I access the dashboard from another smartphone / laptop on the local network (LAN)?</b></summary>

1. Ensure both devices are connected to the same Wi-Fi / LAN network as the host server machine.
2. Find the server machine's IP address (e.g., `192.168.1.50`).
3. Open a browser on the client smartphone/laptop and navigate to:
   ```plaintext
   http://192.168.1.50:8000
   ```
</details>

<details>
<summary><b>❓ Error message: "Excel file is locked" when saving:</b></summary>

Close the Excel spreadsheet `Template_Outage_EIC_Monitoring_unit X.xlsx` if it is currently open in Microsoft Excel desktop so the server can freely write modifications.
</details>

<details>
<summary><b>❓ How do I refresh the data without restarting the browser?</b></summary>

Click the **`⚡ PLTU MSW EIC`** logo badge in the top-left header corner or press **Ctrl + F5** on your keyboard.
</details>

---

<div align="center">

**PLTU MSW &bull; Section Electric, Instrument & Control (EIC)**  
*Continuous Improvement &bull; Excellence in Plant Reliability & Safety*

</div>
