# Product Requirements Document (PRD)
**Project Name:** TIA Portal V20 Universal Diagnostic Reporter  
**Version:** 1.0.0  
**Date:** March 15, 2026  

## 1. Executive Summary
The **TIA Portal V20 Universal Diagnostic Reporter** is a standalone Windows desktop application designed to interface with Siemens TIA Portal V20 via the Openness API. Its primary objective is to perform a deep, recursive audit of a loaded project's hardware configuration to extract and report the diagnostic readiness of every I/O module and channel. 

Specifically, the tool verifies if modules are configured to support `RALRM` (Receive Alarm - OB82) interrupts by checking critical parameters such as "Value Status" (QI), Wire Break, Short Circuit, and Supply Voltage states.

## 2. Problem Statement
In Siemens TIA Portal, configuring I/O modules for hardware diagnostics requires manually enabling specific checkboxes (e.g., Value Status, Wire Break) deep within the hardware properties of each module. Verifying this across a large project is tedious and prone to human error. Furthermore, the Openness API exposes these properties inconsistently depending on the module class:
*   **Standard (ST) Modules:** Diagnostics are stored as module-level attributes.
*   **High Feature (HF) / High Speed (HS) Modules:** Diagnostics are stored on individual channel objects, often with appended suffixes (e.g., `DiagnosticsOverflowForAI`).

This tool abstracts these complexities, actively hunting for diagnostic attributes across all known API paradigms and consolidating them into a single, unified CSV report.

## 3. Target Audience
*   **PLC Programmers / Automation Engineers:** To verify hardware configurations before commissioning or downloading to a physical PLC.
*   **QA / Validation Teams:** To generate compliance reports proving that all safety and diagnostic parameters meet project specifications.

## 4. System Requirements & Technology Stack
*   **Operating System:** Windows 10 / Windows 11 (x64).
*   **Target Framework:** .NET Framework 4.8 (x64 architecture required by TIA Openness).
*   **Language Version:** C# 10.0 (Enabled via `<LangVersion>10.0</LangVersion>` in `.csproj`).
*   **UI Framework:** Windows Forms (WinForms).
*   **Dependencies:** Siemens TIA Portal V20 installed locally (`Siemens.Engineering.dll`).

## 5. Functional Requirements

### 5.1. Connection Management
*   **FR-1.1:** The system shall detect running instances of TIA Portal V20.
*   **FR-1.2:** The system shall attach to the first available TIA Portal process.
*   **FR-1.3:** The system shall securely disconnect and dispose of API COM objects upon user request or application termination.

### 5.2. Hardware Tree Traversal
*   **FR-2.1:** The system shall scan all devices located at the `Root` level of the project.
*   **FR-2.2:** The system shall recursively scan all devices located within nested `DeviceGroups` (User Folders).
*   **FR-2.3:** The system shall scan all devices located within the `UngroupedDevicesGroup`.
*   **FR-2.4:** For every device found, the system shall recursively traverse its `DeviceItems` hierarchy to ensure nested sub-modules and sub-slots are inspected.

### 5.3. Diagnostic Data Extraction
*   **FR-3.1:** The system shall extract attributes without relying on strict, hardcoded attribute names (to support firmware variations like `ForAI`, `ForDQ`).
*   **FR-3.2:** The system shall identify diagnostic attributes by matching substrings: `diag`, `valuestatus`, `wirebreak`, `shortcircuit`, `overflow`, `underflow`, and `nosupplyvoltage`.
*   **FR-3.3 (Module-Level):** The system shall scan the `DeviceItem` directly for diagnostic attributes (targeting ST modules).
*   **FR-3.4 (Channel Property):** The system shall check for a native `Channels` property on the `DeviceItem` and scan each channel object (targeting modern HF modules).
*   **FR-3.5 (Channel Service):** The system shall query the `ChannelProvider` service and scan its channels (targeting integrated I/O on CPUs).

### 5.4. Reporting and Export
*   **FR-4.1:** The system shall provide real-time, threaded UI logging of the scan progress and discovered attributes.
*   **FR-4.2:** The system shall consolidate discovered data into a structured format.
*   **FR-4.3:** The system shall allow the user to export the consolidated data to a `.csv` file.

## 6. Non-Functional Requirements
*   **NFR-1 (Performance):** API calls shall be executed on a background thread (`Task.Run`) to prevent blocking the WinForms UI thread.
*   **NFR-2 (Resilience):** The system shall utilize `dynamic` runtime binding and `try-catch` blocks to prevent the application from crashing when encountering unsupported nodes, read-protected attributes