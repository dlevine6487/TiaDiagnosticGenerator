# Agent Guidelines: TIA Portal V20 Diagnostic Explorer & Generator

## 1. Project Overview
This project is a C# Windows Forms application utilizing the Siemens TIA Portal Openness API (V20). Its dual purpose is to:
1. **Audit & Report:** Scan a TIA Portal project's hardware configuration to extract the diagnostic readiness of every I/O module and channel (e.g., verifying if "Value Status" / "QI", Wire Break, and Short Circuit detection are enabled).
2. **Auto-Generate Code:** Automatically generate the required PLC code (SCL / S7DCL / External Sources) to implement OB82 (`RALRM`) hardware diagnostics natively in the user program, based specifically on Siemens FAQ 109480387.

## 2. Tech Stack & Environment Constraints
*   **Framework:** .NET Framework 4.8 (Requires `<LangVersion>10.0</LangVersion>` in the `.csproj` to support modern C# syntax like nullable reference types).
*   **UI:** Windows Forms (WinForms).
*   **Target API:** TIA Portal Openness V20.
*   **Platform:** x64 architecture (Mandatory for TIA Portal Openness).

## 3. Strict Coding Rules (CRITICAL)

### A. Provide Complete Files (No Brevity)
The user explicitly requires **full, unabbreviated file outputs** for any code changes. Do not use comments like `// ... rest of code ...`. Print the entire `.cs` or `.csproj` file so it can be copied and pasted directly.

### B. .NET 4.8 Compatibility
Do not use .NET Core / .NET 5+ exclusive methods. 
*   **AVOID:** `string.Contains("value", StringComparison.OrdinalIgnoreCase)`
*   **USE:** `string.ToLower().Contains("value")`

### C. Dynamic Binding & Reflection (The "Hw.dll" Bypass)
In some TIA Portal V20 installations, the `Siemens.Engineering.Hw.dll` is either merged into the main assembly or inaccessible. 
*   **DO NOT** use `using Siemens.Engineering.Hw;` or `using Siemens.Engineering.Hw.Features;`.
*   **DO NOT** use explicit typing for hardware components (e.g., `DeviceItem`, `ChannelProvider`).
*   **ALWAYS** use the `dynamic` keyword for traversing the hardware tree (`foreach (dynamic item in items)`).
*   **ALWAYS** use `GetAttribute()` and `GetAttributeInfos()` within `try-catch` blocks to read properties.

### D. Threading and UI Responsiveness
COM calls to TIA Portal are extremely slow.
*   All hardware scanning and file generation must be wrapped in `await Task.Run(() => { ... })`.
*   Because the scan runs on a background thread, **any** updates to the WinForms UI (like logging to a RichTextBox) **must** use `Invoke`:
    ```csharp
    if (this.InvokeRequired) this.Invoke(new Action(() => Log(msg)));
    else { /* Update UI */ }
    ```

## 4. Hardware Traversal Architecture

### A. The Three-Pronged Project Scan
To ensure no device is missed, the scanner must check all three project containment areas:
1.  **Root Devices:** `project.Devices`
2.  **User Groups (Folders):** `project.DeviceGroups` (Requires recursive scanning).
3.  **Ungrouped Devices:** `project.UngroupedDevicesGroup.Devices`

### B. Deep DeviceItem Recursion
I/O cards often hide their diagnostic attributes in child or grandchild `DeviceItems` (sub-modules and sub-slots). 
*   Always recurse through `item.DeviceItems` down to the lowest level.
*   Do not stop recursing just because a node appears to be an empty container.

## 5. Diagnostic Probing Strategy
Siemens hardware categorizes diagnostics into two levels depending on the module type (Standard [ST] vs. High Feature [HF]).

*   **Module-Level Probing (ST Modules):** 
    Attributes like `DiagnosticsWireBreak` exist directly on the module/sub-module node. Iterate over `GetAttributeInfos()`, check if the name contains `"diag"` or `"valuestatus"`, and attempt to read it.
*   **Channel-Level Probing (HF Modules):** 
    Attributes are indexed (e.g., `Channel[0].DiagnosticsWireBreak`). 
*   **Smart Probing (Performance Optimization):** 
    Do not blindly loop through 32 channels. Exception handling in C# COM interop causes massive lag. 
    1. Test `Channel[0].{Attribute}`. 
    2. If it throws an exception, the channel path doesn't exist; skip the rest of the channels immediately. 
    3. If it succeeds, use a `for` loop to increment the channel index until an exception is thrown, marking the end of the module's channels.

## 6. Auto-Generation Strategy (OB82 / RALRM Diagnostics)
Based on Siemens FAQ 109480387, the application must translate discovered hardware into executable PLC code.

### A. Target UDTs (Data Types)
The agent must be able to generate SCL source code to build the following PLC Data Types:
*   `typeDiag`: Master diagnostic struct containing booleans for `error`, `multiError`, `errorCounter`, and an array of `typeChannel` for inputs/outputs.
*   `typeChannel`: Contains an array of `typeChannelDiag` sized to the specific module's channel count (e.g., `[0..15]`).
*   `typeChannelDiag`: Contains `error` (Bool) and `errorCode` (Word).
*   `typeModuleDiag`: Contains specific boolean fault flags (`shortCircuit`, `wireBreak`, `highLimitExceeded`, `lowLimitExceeded`, `noSupplyVoltage`).

### B. Target Code Blocks
*   **FC "ModuleDiag":** A function to decode the `errorCode` returned by `RALRM` into the specific booleans inside `typeModuleDiag`.
*   **FB "1x00Diag82":** The master Function Block that calls the `RALRM` instruction (MODE := 2). It maps `AINFO` and `TINFO` parameter data into the UDTs for individual channels based on the Hardware Identifier (`HW_IO`).

### C. Generation Execution
1.  During the hardware audit, cache the `Name`, `ChannelCount`, and `HardwareIdentifier` (`HW_IO` system constant) of every module that is found to support diagnostics.
2.  Generate an SCL External Source file containing the UDTs, FC, and FB.
3.  Generate an instance call inside a generic OB82 wrapper block for *every* discovered I/O card, automatically mapping its `HW_IO` parameter to the `F_ID` pin of the `1x00Diag82` block.