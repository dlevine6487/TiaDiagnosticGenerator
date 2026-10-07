# TIA Diagnostic Reader - Execution Instructions

The `TiaDiagnosticReader` is a Windows Forms application written in C# targeting `.NET Framework 4.8`. It utilizes the TIA Portal Openness API to connect to running TIA Portal instances, scan the hardware topology, and read diagnostic attributes.

## Prerequisites

1. **Operating System:** Windows (required for TIA Portal Openness).
2. **TIA Portal:** TIA Portal V20 (or V21) must be installed.
   - The application relies on `Siemens.Engineering.dll` located at `C:\Program Files\Siemens\Automation\Portal V20\PublicAPI\V20\`.
   - *Note: If you are using a different version of TIA Portal (e.g., V21), you must update the `<HintPath>` in `Version 1/TiaDiagnosticReader/TiaDiagnosticReader.csproj` to point to the correct `PublicAPI` folder.*
3. **Openness Authorization:** You must be a member of the local Windows user group `Siemens TIA Openness` to interact with the API.
4. **Development Environment:** Visual Studio 2022 (recommended) or the .NET SDK.

## Compiling and Running

1. **Open the Solution:**
   Navigate to the `Version 1` folder and open the `Version 1.sln` file in Visual Studio.

2. **Build the Project:**
   Build the project in Visual Studio (`Ctrl + Shift + B`). Alternatively, if you are using the .NET CLI:
   ```shell
   cd "Version 1/TiaDiagnosticReader"
   dotnet build
   ```

3. **Start TIA Portal:**
   Before running the application, ensure that an instance of TIA Portal is actively running and a project is open. The tool will attempt to attach to the first running instance it finds.

4. **Run the Application:**
   Start the application (`F5` in Visual Studio).
   - Click **"Connect & Scan All"** to attach to TIA Portal and extract the hardware diagnostics.
   - Once the scan finishes, you can view the logs in the styled RichTextBox or click **"Export to CSV"** to save the findings.

## UI / Theme Updates

The UI has been enhanced to align with the **Siemens iX (Industrial Experience) Design System**.
- **Dark Theme:** The application utilizes a dark background (`#121419` / `#1B1E23`).
- **Typography:** The standard font has been updated to `Segoe UI` for a cleaner, modern look.
- **Accents:** Buttons are flat, using the primary Siemens teal (`#00646E` / `#008A94`) and secondary grays (`#646E78`), removing the legacy 3D borders for a modern industrial interface.

## Troubleshooting

- **CS0246 (Missing Siemens.Engineering):** If the build fails with a missing `Siemens.Engineering` reference, verify the absolute path in `TiaDiagnosticReader.csproj` matches your local TIA Portal installation path.
- **Connection Refused:** Ensure TIA Portal is running, a project is fully loaded, and you have clicked "Yes" on the Openness Firewall prompt if it appears.
