# Product Requirements Document (PRD)
**Project Name:** TIA Portal V20 Diagnostic Explorer & Generator  
**Version:** 2.0.0  
**Date:** March 15, 2026  

## 1. Executive Summary
The **TIA Portal V20 Diagnostic Explorer & Generator** is a standalone Windows desktop application designed to interface with Siemens TIA Portal V20 via the Openness API. Its objective is two-fold:
1. **Audit:** Perform a deep, recursive audit of a loaded project's hardware configuration to extract and report the diagnostic readiness of every I/O module and individual channel.
2. **Generate:** Automatically construct the SCL/S7DCL program logic required to read, decode, and report these hardware faults inside the PLC user program, adhering strictly to the Siemens standard for `RALRM` (OB82) module and channel diagnostics (FAQ 109480387).

## 2. Problem Statement
Configuring I/O modules for hardware diagnostics in TIA Portal requires manually enabling checkboxes (e.g., Value Status, Wire Break) deep within the hardware properties. Once enabled, the PLC programmer must manually write tedious, repetitive boilerplate code (UDTs, OB82 triggers, `RALRM` calls, and decoding logic) for *every single module and channel* to make those faults visible to the SCADA/HMI. 

This tool abstracts these complexities by actively hunting for configured diagnostic attributes and immediately generating the exact, bespoke PLC code needed to monitor them on a per-channel basis.

## 3. Target Audience
*   **PLC Programmers / Automation Engineers:** To instantly verify hardware configurations and eliminate the hours spent writing boilerplate OB82 diagnostic mapping code.
*   **QA / Validation Teams:** To generate compliance reports proving that all safety and diagnostic parameters are actively monitored.

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

### 5.3. Diagnostic Data Extraction (Audit)
*   **FR-3.1:** The system shall identify diagnostic attributes by matching substrings: `diag`, `valuestatus`, `wirebreak`, `shortcircuit`, `overflow`, `underflow`, and `nosupplyvoltage`.
*   **FR-3.2 (Module-Level):** The system shall scan the `DeviceItem` directly for diagnostic attributes (targeting ST modules).
*   **FR-3.3 (Channel Property):** The system shall check for a native `Channels` property on the `DeviceItem` and scan each channel object (targeting modern HF modules).
*   **FR-3.4 (Channel Service):** The system shall query the `ChannelProvider` service and scan its channels (targeting integrated I/O on CPUs).
*   **FR-3.5:** The system shall log the capability type (Module-Only Diagnostics vs Channel-Level Diagnostics) by parsing the module's `TypeIdentifier`.

### 5.4. Auto-Generation of PLC Diagnostic Logic (Code Generation)
*   **FR-4.1:** The system shall generate PLC Data Types (UDTs) to hold diagnostic states. These UDTs (`typeDiag`, `typeChannelType`, `typeChannelDiag`, `typeModuleDiag`) must conform to the structures outlined in Siemens FAQ 109480387.
*   **FR-4.2:** The system shall dynamically size the arrays within the generated UDTs based on the `ChannelCount` scraped from the actual physical modules discovered in the audit.
*   **FR-4.3:** The system shall generate a master Function Block (e.g., `FB 1x00Diag82`) that utilizes the `RALRM` (Receive Alarm) instruction.
*   **FR-4.4:** The system shall generate a decoding Function (e.g., `FC ModuleDiag`) to map hex `errorCode` values to distinct Boolean fault flags for individual channels.
*   **FR-4.5:** The system shall generate OB82 (Diagnostic Error Interrupt) logic that instantiates the master FB for every configured module, passing the exact Hardware Identifier (`HW_IO`) of the discovered module into the block call.
*   **FR-4.6:** The system shall output the generated logic as structured text (.SCL) or XML files ready for import into the TIA Portal Project.

### 5.5. Reporting and Export
*   **FR-5.1:** The system shall provide real-time, threaded UI logging of the scan progress.
*   **FR-5.2:** The system shall consolidate discovered hardware audit data into a structured CSV format and allow the user to export it.

## 6. Non-Functional Requirements
*   **NFR-1 (Performance):** API calls shall be executed on a background thread (`Task.Run`) to prevent blocking the WinForms UI thread.
*   **NFR-2 (Resilience):** The system shall utilize `dynamic` runtime binding and `try-catch` blocks to prevent the application from crashing when encountering unsupported nodes, read-protected attributes, or missing `Hw` assemblies in merged TIA V20 installations.
*   **NFR-3 (Data Integrity):** The system shall use a `HashSet` to prevent duplicate diagnostic entries if multiple extraction methods yield the same attribute.

## 7. Data Dictionary (CSV Export Format)
The exported CSV file shall contain the following columns:

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **Station** | The parent PLC or distributed I/O station name. | `S71500/ET200MP station_1` |
| **Item** | The specific name of the module or sub-module. | `AQ 4xU/I ST_1` |
| **Type** | The Siemens Order Number / Type Identifier. | `OrderNumber:6ES7 135-6HD00-0BA1/V1.1` |
| **Location** | Where the attribute was found (`Module` or `Ch_X`). | `Channel[0]` |
| **Attribute** | The exact internal API name of the diagnostic property. | `DiagnosticsWireBreak` |
| **Value** | The configured value in TIA Portal (usually True/False). | `True` |

## 8. Known Limitations & Constraints
*   **F-Modules (Failsafe):** Safety I/O modules may hide their diagnostic configurations behind protected safety services. These may not appear in standard attribute dumps.
*   **Hardware Changes:** If the hardware configuration in TIA Portal changes, the scan must be re-run, and the generated SCL must be re-imported to update channel mapping bounds.