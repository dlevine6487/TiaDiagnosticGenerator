using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using Siemens.Engineering;

namespace TiaDiagnosticGui
{
    public partial class Form1 : Form
    {
        private TiaPortal? tiaPortal;
        private Project? project;
        private Button btnConnect;
        private Button btnDisconnect;
        private Button btnExportCsv;
        private Button btnGenerateScl;
        private RichTextBox txtOutput;
        private List<string> csvReportData;
        private List<DiagnosticModuleInfo> diagnosticModules;

        public class DiagnosticModuleInfo
        {
            public string ModuleName { get; set; } = "";
            public string HardwareIdentifier { get; set; } = "";
            public int ChannelCount { get; set; } = 0;
            public bool IsHighFeature { get; set; } = false;
        }

        public Form1()
        {
            this.Text = "TIA Portal V20 Ultimate Diagnostic Scanner";
            this.Size = new System.Drawing.Size(1000, 800);
            this.csvReportData = new List<string>();
            this.diagnosticModules = new List<DiagnosticModuleInfo>();

            btnConnect = new Button();
            btnConnect.Text = "Connect & Scan All";
            btnConnect.Location = new System.Drawing.Point(10, 10);
            btnConnect.Width = 150;
            btnConnect.Click += BtnConnect_Click;

            btnDisconnect = new Button();
            btnDisconnect.Text = "Disconnect";
            btnDisconnect.Location = new System.Drawing.Point(170, 10);
            btnDisconnect.Width = 100;
            btnDisconnect.Enabled = false;
            btnDisconnect.Click += BtnDisconnect_Click;

            btnExportCsv = new Button();
            btnExportCsv.Text = "Export to CSV";
            btnExportCsv.Location = new System.Drawing.Point(280, 10);
            btnExportCsv.Width = 120;
            btnExportCsv.Enabled = false;
            btnExportCsv.Click += BtnExportCsv_Click;

            btnGenerateScl = new Button();
            btnGenerateScl.Text = "Generate SCL";
            btnGenerateScl.Location = new System.Drawing.Point(410, 10);
            btnGenerateScl.Width = 120;
            btnGenerateScl.Enabled = false;
            btnGenerateScl.Click += BtnGenerateScl_Click;

            txtOutput = new RichTextBox();
            txtOutput.Location = new System.Drawing.Point(10, 45);
            txtOutput.Width = 960;
            txtOutput.Height = 700;
            txtOutput.ReadOnly = true;
            txtOutput.ScrollBars = RichTextBoxScrollBars.Vertical;
            txtOutput.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtOutput.Font = new System.Drawing.Font("Consolas", 8);

            this.Controls.Add(btnConnect);
            this.Controls.Add(btnDisconnect);
            this.Controls.Add(btnExportCsv);
            this.Controls.Add(btnGenerateScl);
            this.Controls.Add(txtOutput);
        }

        private async void BtnGenerateScl_Click(object? sender, EventArgs e)
        {
            if (diagnosticModules.Count == 0)
            {
                Log("No modules with diagnostics were found to generate SCL for.");
                return;
            }

            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select a folder to save the generated S7DCL/S7res files";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    btnGenerateScl.Enabled = false;
                    Log("Generating SCL/S7DCL files...");

                    await Task.Run(() =>
                    {
                        try
                        {
                            GenerateAndSaveVciFiles(fbd.SelectedPath);
                            Log($"\n[SCL GENERATION] Successfully saved VCI files to {fbd.SelectedPath}");
                        }
                        catch (Exception ex)
                        {
                            Log($"Error generating files: {ex.Message}");
                        }
                    });

                    btnGenerateScl.Enabled = true;
                }
            }
        }

        private void GenerateAndSaveVciFiles(string outDir)
        {
            string udtDir = Path.Combine(outDir, "UDT");
            Directory.CreateDirectory(udtDir);

            // Find max channel count
            int maxChannelsAll = 1;
            foreach (var mod in diagnosticModules)
            {
                if (mod.ChannelCount > maxChannelsAll) maxChannelsAll = mod.ChannelCount;
            }

            // ==========================================
            // Generate UDTs (S7DCL & S7res)
            // ==========================================

            // 1. typeModuleDiag
            string tmDiagS7dcl = @"    TYPE
        typeModuleDiag : STRUCT
        { S7_MLC := ""MLC_modSc"" }
        shortCircuit : Bool;
        { S7_MLC := ""MLC_modWb"" }
        wireBreak : Bool;
        { S7_MLC := ""MLC_modHle"" }
        highLimitExceeded : Bool;
        { S7_MLC := ""MLC_modLle"" }
        lowLimitExceeded : Bool;
        { S7_MLC := ""MLC_modNsv"" }
        noSupplyVoltage : Bool;
        END_STRUCT;
    END_TYPE";

            string tmDiagS7res = @"MultiLingualTexts:
  - id: MLC_modSc
    en-US: Short circuit
  - id: MLC_modWb
    en-US: Wire break
  - id: MLC_modHle
    en-US: High limit exceeded
  - id: MLC_modLle
    en-US: Low limit exceeded
  - id: MLC_modNsv
    en-US: No supply voltage";

            File.WriteAllText(Path.Combine(udtDir, "typeModuleDiag.s7dcl"), tmDiagS7dcl);
            File.WriteAllText(Path.Combine(udtDir, "typeModuleDiag.s7res"), tmDiagS7res);

            // 2. typeChannelDiag
            string tcDiagS7dcl = @"    TYPE
        typeChannelDiag : STRUCT
        { S7_MLC := ""MLC_chErr"" }
        error : Bool;
        { S7_MLC := ""MLC_chErrCode"" }
        errorCode : Word;
        END_STRUCT;
    END_TYPE";

            string tcDiagS7res = @"MultiLingualTexts:
  - id: MLC_chErr
    en-US: Error existing
  - id: MLC_chErrCode
    en-US: Error code";

            File.WriteAllText(Path.Combine(udtDir, "typeChannelDiag.s7dcl"), tcDiagS7dcl);
            File.WriteAllText(Path.Combine(udtDir, "typeChannelDiag.s7res"), tcDiagS7res);

            // 3. typeChannel
            string tChanS7dcl = $@"    TYPE
        typeChannel : STRUCT
        {{ S7_MLC := ""MLC_chArr"" }}
        channel : Array[0..{maxChannelsAll - 1}] of ""typeChannelDiag"";
        END_STRUCT;
    END_TYPE";

            string tChanS7res = @"MultiLingualTexts:
  - id: MLC_chArr
    en-US: Array of channel diagnostics";

            File.WriteAllText(Path.Combine(udtDir, "typeChannel.s7dcl"), tChanS7dcl);
            File.WriteAllText(Path.Combine(udtDir, "typeChannel.s7res"), tChanS7res);

            // 4. typeDiag
            string tDiagS7dcl = @"    TYPE
        typeDiag : STRUCT
        { S7_MLC := ""MLC_dgErr"" }
        error : Bool;
        { S7_MLC := ""MLC_dgMul"" }
        multiError : Bool;
        { S7_MLC := ""MLC_dgCnt"" }
        errorCounter : Int;
        { S7_MLC := ""MLC_dgCh"" }
        channels : ""typeChannel"";
        END_STRUCT;
    END_TYPE";

            string tDiagS7res = @"MultiLingualTexts:
  - id: MLC_dgErr
    en-US: Overall Error
  - id: MLC_dgMul
    en-US: Multiple Errors Present
  - id: MLC_dgCnt
    en-US: Error Counter
  - id: MLC_dgCh
    en-US: Channels";

            File.WriteAllText(Path.Combine(udtDir, "typeDiag.s7dcl"), tDiagS7dcl);
            File.WriteAllText(Path.Combine(udtDir, "typeDiag.s7res"), tDiagS7res);


            // ==========================================
            // Generate Code Blocks (.scl for SCL blocks)
            // ==========================================

            // 5. FC ModuleDiag
            System.Text.StringBuilder fcScl = new System.Text.StringBuilder();
            fcScl.AppendLine("FUNCTION \"ModuleDiag\" : Void");
            fcScl.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            fcScl.AppendLine("VERSION : 0.1");
            fcScl.AppendLine("   VAR_INPUT");
            fcScl.AppendLine("      errorCode : Word;");
            fcScl.AppendLine("   END_VAR");
            fcScl.AppendLine("   VAR_IN_OUT");
            fcScl.AppendLine("      moduleDiag : \"typeModuleDiag\";");
            fcScl.AppendLine("   END_VAR");
            fcScl.AppendLine("BEGIN");
            fcScl.AppendLine("   // Reset all errors first");
            fcScl.AppendLine("   #moduleDiag.shortCircuit := FALSE;");
            fcScl.AppendLine("   #moduleDiag.wireBreak := FALSE;");
            fcScl.AppendLine("   #moduleDiag.highLimitExceeded := FALSE;");
            fcScl.AppendLine("   #moduleDiag.lowLimitExceeded := FALSE;");
            fcScl.AppendLine("   #moduleDiag.noSupplyVoltage := FALSE;");
            fcScl.AppendLine("");
            fcScl.AppendLine("   CASE WORD_TO_INT(#errorCode) OF");
            fcScl.AppendLine("      1:");
            fcScl.AppendLine("         #moduleDiag.shortCircuit := TRUE;");
            fcScl.AppendLine("      6:");
            fcScl.AppendLine("         #moduleDiag.wireBreak := TRUE;");
            fcScl.AppendLine("      7:");
            fcScl.AppendLine("         #moduleDiag.highLimitExceeded := TRUE;");
            fcScl.AppendLine("      8:");
            fcScl.AppendLine("         #moduleDiag.lowLimitExceeded := TRUE;");
            fcScl.AppendLine("      17:");
            fcScl.AppendLine("         #moduleDiag.noSupplyVoltage := TRUE;");
            fcScl.AppendLine("   END_CASE;");
            fcScl.AppendLine("END_FUNCTION");
            File.WriteAllText(Path.Combine(outDir, "ModuleDiag.scl"), fcScl.ToString());

            // 6. FB 1x00Diag82
            System.Text.StringBuilder fbScl = new System.Text.StringBuilder();
            fbScl.AppendLine("FUNCTION_BLOCK \"1x00Diag82\"");
            fbScl.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            fbScl.AppendLine("VERSION : 0.1");
            fbScl.AppendLine("   VAR_INPUT");
            fbScl.AppendLine("      F_ID : HW_IO;");
            fbScl.AppendLine("   END_VAR");
            fbScl.AppendLine("   VAR_OUTPUT");
            fbScl.AppendLine("      new : Bool;");
            fbScl.AppendLine("      status : DWord;");
            fbScl.AppendLine("      id : HW_IO;");
            fbScl.AppendLine("      len : UInt;");
            fbScl.AppendLine("      areaLenError : Bool;");
            fbScl.AppendLine("   END_VAR");
            fbScl.AppendLine("   VAR_IN_OUT");
            fbScl.AppendLine("      Diag : \"typeDiag\";");
            fbScl.AppendLine("   END_VAR");
            fbScl.AppendLine("   VAR");
            fbScl.AppendLine("      ralrm : RALRM;");
            fbScl.AppendLine("      tinfo : TINFO_OB;");
            fbScl.AppendLine("      ainfo : AINFO_OB;");
            fbScl.AppendLine("      channelNumber : Int;");
            fbScl.AppendLine("      channelError : Bool;");
            fbScl.AppendLine("   END_VAR");
            fbScl.AppendLine("BEGIN");
            fbScl.AppendLine("   #ralrm(MODE := 2,");
            fbScl.AppendLine("          F_ID := #F_ID,");
            fbScl.AppendLine("          NEW => #new,");
            fbScl.AppendLine("          STATUS => #status,");
            fbScl.AppendLine("          ID => #id,");
            fbScl.AppendLine("          LEN => #len,");
            fbScl.AppendLine("          TINFO := #tinfo,");
            fbScl.AppendLine("          AINFO := #ainfo);");
            fbScl.AppendLine("");
            fbScl.AppendLine("   IF #new THEN");
            fbScl.AppendLine("      #channelNumber := UINT_TO_INT(#ainfo.USI); // Channel number from User Structure Identifier");
            fbScl.AppendLine("      #channelError := (#tinfo.OB_NUM = 82) AND (#tinfo.PRG_INFO = 16#39); // Event incoming");
            fbScl.AppendLine("");
            fbScl.AppendLine("      // Ensure channel is within array bounds");
            fbScl.AppendLine($"      IF (#channelNumber >= 0) AND (#channelNumber <= {maxChannelsAll - 1}) THEN");
            fbScl.AppendLine("         #Diag.channels.channel[#channelNumber].error := #channelError;");
            fbScl.AppendLine("         IF #channelError THEN");
            fbScl.AppendLine("            #Diag.channels.channel[#channelNumber].errorCode := #ainfo.ERR_MOD; // Example, map to correct error word");
            fbScl.AppendLine("         ELSE");
            fbScl.AppendLine("            #Diag.channels.channel[#channelNumber].errorCode := 16#0;");
            fbScl.AppendLine("         END_IF;");
            fbScl.AppendLine("      END_IF;");
            fbScl.AppendLine("");
            fbScl.AppendLine("      // Multi-error handling (simplified)");
            fbScl.AppendLine("      IF #channelError THEN");
            fbScl.AppendLine("         #Diag.errorCounter := #Diag.errorCounter + 1;");
            fbScl.AppendLine("      ELSIF #Diag.errorCounter > 0 THEN");
            fbScl.AppendLine("         #Diag.errorCounter := #Diag.errorCounter - 1;");
            fbScl.AppendLine("      END_IF;");
            fbScl.AppendLine("");
            fbScl.AppendLine("      #Diag.error := #Diag.errorCounter > 0;");
            fbScl.AppendLine("      #Diag.multiError := #Diag.errorCounter > 1;");
            fbScl.AppendLine("   END_IF;");
            fbScl.AppendLine("END_FUNCTION_BLOCK");
            File.WriteAllText(Path.Combine(outDir, "1x00Diag82.scl"), fbScl.ToString());

            // 7. Data block for tags (Tags.s7dcl)
            System.Text.StringBuilder tagsS7dcl = new System.Text.StringBuilder();
            tagsS7dcl.AppendLine(@"        {
           S7_Optimized := ""TRUE"";
           S7_StandardRetain := ""FALSE"";
           S7_Version := ""0.1""
        }");
            tagsS7dcl.AppendLine("    DATA_BLOCK Tags");
            tagsS7dcl.AppendLine("        VAR");

            // Generate array for output statuses
            if (diagnosticModules.Count > 0)
            {
                tagsS7dcl.AppendLine($"            diag82 : Array[1..{diagnosticModules.Count}] of \"typeDiag82\";");
            }

            int idx = 1;
            foreach (var mod in diagnosticModules)
            {
                // Generate safe names without spaces or invalid chars
                string safeName = System.Text.RegularExpressions.Regex.Replace(mod.ModuleName, @"[^a-zA-Z0-9_]", "");
                if (string.IsNullOrEmpty(safeName) || char.IsDigit(safeName[0]))
                {
                    safeName = "Mod" + safeName + "_" + idx;
                }
                tagsS7dcl.AppendLine($"            {safeName} : \"typeDiag\";");
                idx++;
            }
            tagsS7dcl.AppendLine("        END_VAR");
            tagsS7dcl.AppendLine("    END_DATA_BLOCK");
            File.WriteAllText(Path.Combine(outDir, "Tags.s7dcl"), tagsS7dcl.ToString());

            // Add typeDiag82 struct to match the example outputs
            string tDiag82S7dcl = @"    TYPE
        typeDiag82 : STRUCT
        new : Bool;
        status : DWord;
        id : HW_IO;
        len : UInt;
        areaLenError : Bool;
        END_STRUCT;
    END_TYPE";
            File.WriteAllText(Path.Combine(udtDir, "typeDiag82.s7dcl"), tDiag82S7dcl);

            // 8. OB82 (DiagnosticErrorInterrupt) in XML format (FBD representation)
            System.Text.StringBuilder obXml = new System.Text.StringBuilder();
            obXml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            obXml.AppendLine("<Document>");
            obXml.AppendLine("  <Engineering version=\"V20\" />");
            obXml.AppendLine("  <SW.Blocks.OB ID=\"0\">");
            obXml.AppendLine("    <AttributeList>");
            obXml.AppendLine("      <Interface><Sections xmlns=\"http://www.siemens.com/automation/Openness/SW/Interface/v5\">");
            obXml.AppendLine("  <Section Name=\"Input\">");
            obXml.AppendLine("    <Member Name=\"IO_State\" Datatype=\"Word\" Informative=\"true\"></Member>");
            obXml.AppendLine("    <Member Name=\"LADDR\" Datatype=\"HW_ANY\" Informative=\"true\"></Member>");
            obXml.AppendLine("    <Member Name=\"Channel\" Datatype=\"UInt\" Informative=\"true\"></Member>");
            obXml.AppendLine("    <Member Name=\"MultiError\" Datatype=\"Bool\" Informative=\"true\"></Member>");
            obXml.AppendLine("  </Section>");
            obXml.AppendLine("  <Section Name=\"Temp\" />");
            obXml.AppendLine("  <Section Name=\"Constant\" />");
            obXml.AppendLine("</Sections></Interface>");
            obXml.AppendLine("      <MemoryLayout>Optimized</MemoryLayout>");
            obXml.AppendLine("      <Name>DiagnosticErrorInterrupt</Name>");
            obXml.AppendLine("      <Namespace />");
            obXml.AppendLine("      <Number>82</Number>");
            obXml.AppendLine("      <ProgrammingLanguage>FBD</ProgrammingLanguage>");
            obXml.AppendLine("      <SecondaryType>DiagnosticErrorInterrupt</SecondaryType>");
            obXml.AppendLine("      <SetENOAutomatically>false</SetENOAutomatically>");
            obXml.AppendLine("    </AttributeList>");
            obXml.AppendLine("    <ObjectList>");
            obXml.AppendLine("      <SW.Blocks.CompileUnit ID=\"3\" CompositionName=\"CompileUnits\">");
            obXml.AppendLine("        <AttributeList>");
            obXml.AppendLine("          <NetworkSource><FlgNet xmlns=\"http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v5\">");
            obXml.AppendLine("  <Parts>");

            idx = 1;
            int uidBase = 20; // Ensure unique UIds per part

            foreach (var mod in diagnosticModules)
            {
                string safeName = System.Text.RegularExpressions.Regex.Replace(mod.ModuleName, @"[^a-zA-Z0-9_]", "");
                if (string.IsNullOrEmpty(safeName) || char.IsDigit(safeName[0])) safeName = "Mod" + safeName + "_" + idx;

                string hwIdStr = string.IsNullOrEmpty(mod.HardwareIdentifier) ? "0" : mod.HardwareIdentifier;

                // Hardware constant access
                obXml.AppendLine($"    <Access Scope=\"GlobalConstant\" UId=\"{uidBase + 1}\">");
                obXml.AppendLine($"      <Constant Name=\"{hwIdStr}\" />"); // E.g., Local~AI_5_AQ_2_1
                obXml.AppendLine("    </Access>");

                // Diag struct access
                obXml.AppendLine($"    <Access Scope=\"GlobalVariable\" UId=\"{uidBase + 2}\">");
                obXml.AppendLine("      <Symbol>");
                obXml.AppendLine("        <Component Name=\"Tags\" />");
                obXml.AppendLine($"        <Component Name=\"{safeName}\" />");
                obXml.AppendLine("      </Symbol>");
                obXml.AppendLine("    </Access>");

                // diag82 variables access
                string[] outVars = { "new", "status", "id", "len", "areaLenError" };
                for (int i = 0; i < outVars.Length; i++)
                {
                    obXml.AppendLine($"    <Access Scope=\"GlobalVariable\" UId=\"{uidBase + 10 + i}\">");
                    obXml.AppendLine("      <Symbol>");
                    obXml.AppendLine("        <Component Name=\"Tags\" />");
                    obXml.AppendLine("        <Component Name=\"diag82\" AccessModifier=\"Array\">");
                    obXml.AppendLine("          <Access Scope=\"LiteralConstant\">");
                    obXml.AppendLine("            <Constant>");
                    obXml.AppendLine("              <ConstantType>DInt</ConstantType>");
                    obXml.AppendLine($"              <ConstantValue>{idx}</ConstantValue>");
                    obXml.AppendLine("            </Constant>");
                    obXml.AppendLine("          </Access>");
                    obXml.AppendLine("        </Component>");
                    obXml.AppendLine($"        <Component Name=\"{outVars[i]}\" />");
                    obXml.AppendLine("      </Symbol>");
                    obXml.AppendLine("    </Access>");
                }

                // FB Call
                obXml.AppendLine($"    <Call UId=\"{uidBase + 3}\">");
                obXml.AppendLine("      <CallInfo Name=\"1x00Diag82\" BlockType=\"FB\">");
                obXml.AppendLine($"        <Instance Scope=\"GlobalVariable\" UId=\"{uidBase + 4}\">");
                obXml.AppendLine($"          <Component Name=\"Inst_{safeName}\" />");
                obXml.AppendLine("        </Instance>");
                obXml.AppendLine("        <Parameter Name=\"fId\" Section=\"Input\" Type=\"HW_IO\" />");
                obXml.AppendLine("        <Parameter Name=\"new\" Section=\"Output\" Type=\"Bool\" />");
                obXml.AppendLine("        <Parameter Name=\"status\" Section=\"Output\" Type=\"DWord\" />");
                obXml.AppendLine("        <Parameter Name=\"id\" Section=\"Output\" Type=\"HW_IO\" />");
                obXml.AppendLine("        <Parameter Name=\"len\" Section=\"Output\" Type=\"UInt\" />");
                obXml.AppendLine("        <Parameter Name=\"areaLenError\" Section=\"Output\" Type=\"Bool\" />");
                obXml.AppendLine("        <Parameter Name=\"diag\" Section=\"InOut\" Type=\"&quot;typeDiag&quot;\" />");
                obXml.AppendLine("      </CallInfo>");
                obXml.AppendLine("    </Call>");

                uidBase += 20;
                idx++;

                // Create instance DB (s7dcl) for each card
                string instDb = $@"        {{
           S7_Optimized := ""TRUE"";
           S7_StandardRetain := ""FALSE"";
           S7_Version := ""0.1""
        }}
    DATA_BLOCK Inst_{safeName} : ""1x00Diag82""
    END_DATA_BLOCK";
                File.WriteAllText(Path.Combine(outDir, $"Inst_{safeName}.s7dcl"), instDb);
            }

            obXml.AppendLine("  </Parts>");
            obXml.AppendLine("  <Wires>");

            idx = 1;
            uidBase = 20;
            foreach (var mod in diagnosticModules)
            {
                obXml.AppendLine($"    <Wire UId=\"{uidBase + 5}\">");
                obXml.AppendLine($"      <OpenCon UId=\"{uidBase + 6}\" />");
                obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"en\" />");
                obXml.AppendLine("    </Wire>");
                obXml.AppendLine($"    <Wire UId=\"{uidBase + 7}\">");
                obXml.AppendLine($"      <IdentCon UId=\"{uidBase + 1}\" />");
                obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"fId\" />");
                obXml.AppendLine("    </Wire>");

                string[] outVars = { "new", "status", "id", "len", "areaLenError" };
                for (int i = 0; i < outVars.Length; i++)
                {
                    obXml.AppendLine($"    <Wire UId=\"{uidBase + 20 + i}\">");
                    obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"{outVars[i]}\" />");
                    obXml.AppendLine($"      <IdentCon UId=\"{uidBase + 10 + i}\" />");
                    obXml.AppendLine("    </Wire>");
                }

                obXml.AppendLine($"    <Wire UId=\"{uidBase + 8}\">");
                obXml.AppendLine($"      <IdentCon UId=\"{uidBase + 2}\" />");
                obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"diag\" />");
                obXml.AppendLine("    </Wire>");
                uidBase += 20;
                idx++;
            }

            obXml.AppendLine("  </Wires>");
            obXml.AppendLine("</FlgNet></NetworkSource>");
            obXml.AppendLine("          <ProgrammingLanguage>FBD</ProgrammingLanguage>");
            obXml.AppendLine("        </AttributeList>");
            obXml.AppendLine("      </SW.Blocks.CompileUnit>");
            obXml.AppendLine("    </ObjectList>");
            obXml.AppendLine("  </SW.Blocks.OB>");
            obXml.AppendLine("</Document>");
            File.WriteAllText(Path.Combine(outDir, "DiagnosticErrorInterrupt.xml"), obXml.ToString());
        }

        private async void BtnConnect_Click(object? sender, EventArgs e)
        {
            txtOutput.Clear();
            csvReportData.Clear();
            csvReportData.Add("Station,Item,Type,Location,Attribute,Value");
            diagnosticModules.Clear();

            btnConnect.Enabled = false;
            Log("Initializing TIA Portal connection on background thread...");

            await Task.Run(() => PerformScan());

            btnDisconnect.Enabled = (tiaPortal != null);
            btnExportCsv.Enabled = (csvReportData.Count > 1);
            btnGenerateScl.Enabled = (diagnosticModules.Count > 0);
            btnConnect.Enabled = (tiaPortal == null);
        }

        private void PerformScan()
        {
            try
            {
                var processes = TiaPortal.GetProcesses();
                if (processes.Count == 0)
                {
                    Log("Error: No running TIA Portal instances found.");
                    return;
                }

                tiaPortal = processes[0].Attach();
                project = tiaPortal.Projects[0];

                Log($"Connected to: {project.Name}");

                foreach (dynamic device in project.Devices) WalkDevice(device);
                ScanGroupsRecursive(project.DeviceGroups);
                if (project.UngroupedDevicesGroup != null)
                {
                    foreach (dynamic device in project.UngroupedDevicesGroup.Devices) WalkDevice(device);
                }

                Log("\n[SCAN] Completed successfully.");
            }
            catch (Exception ex)
            {
                Log($"Critical Error: {ex.Message}");
            }
        }

        private void ScanGroupsRecursive(dynamic groups)
        {
            foreach (dynamic group in groups)
            {
                foreach (dynamic device in group.Devices) WalkDevice(device);
                if (group.Groups != null) ScanGroupsRecursive(group.Groups);
            }
        }

        private void WalkDevice(dynamic device)
        {
            if (device == null) return;
            string stationName = "Unknown";
            try { stationName = device.Name.ToString(); } catch { }

            Log($"\n>>> STATION: {stationName}");
            RecursiveWalk(device.DeviceItems, stationName);
        }

        private void RecursiveWalk(dynamic items, string stationName)
        {
            if (items == null) return;
            foreach (dynamic item in items)
            {
                if (item != null)
                {
                    ProbeAllDiagnostics(item, stationName);

                    // Recursively process children to ensure we hit sub-modules
                    if (item.DeviceItems != null)
                    {
                        RecursiveWalk(item.DeviceItems, stationName);
                    }
                }
            }
        }

        private void ProbeAllDiagnostics(dynamic item, string stationName)
        {
            string itemName = "Unknown";
            string typeId = "Unknown";
            string hwId = "0";

            try { itemName = item.Name.ToString(); } catch { }
            try { typeId = item.TypeIdentifier.ToString(); } catch { }

            // Attempt to get HW_IO identifier
            try
            {
                dynamic attrs = item.GetAttributeInfos();
                if (attrs != null)
                {
                    foreach (var attr in attrs)
                    {
                        if (attr.Name.ToString().Equals("HardwareIdentifier", StringComparison.OrdinalIgnoreCase))
                        {
                            hwId = item.GetAttribute("HardwareIdentifier").ToString();
                            break;
                        }
                    }
                }
            }
            catch { }

            // Keywords we are looking for in the attribute names
            string[] keywords = { "diag", "valuestatus", "wirebreak", "shortcircuit", "overflow", "underflow", "nosupplyvoltage" };
            
            // Track seen attributes to prevent duplicate logging if multiple methods find the same data
            HashSet<string> seenAttrs = new HashSet<string>();
            bool hasDiagnostics = false;
            int maxChannelCount = 0;

            // Universal Attribute Extractor Method
            Action<dynamic, string, string> checkAndLogAttrs = (targetObj, logPrefix, location) =>
            {
                try
                {
                    var attrs = targetObj.GetAttributeInfos();
                    if (attrs != null)
                    {
                        foreach (var attr in attrs)
                        {
                            string nameLower = attr.Name.ToLower();
                            bool isDiagnostic = false;
                            
                            foreach (var kw in keywords)
                            {
                                if (nameLower.Contains(kw))
                                {
                                    isDiagnostic = true;
                                    break;
                                }
                            }

                            if (isDiagnostic)
                            {
                                try
                                {
                                    object val = targetObj.GetAttribute(attr.Name);
                                    if (val != null)
                                    {
                                        hasDiagnostics = true;

                                        // Update channel count based on location string
                                        if (location.StartsWith("Ch_"))
                                        {
                                            if (int.TryParse(location.Substring(3), out int chIdx))
                                            {
                                                maxChannelCount = Math.Max(maxChannelCount, chIdx + 1);
                                            }
                                        }
                                        else if (location == "Module" && maxChannelCount == 0)
                                        {
                                            // Fallback for ST modules that don't have explicit channel properties
                                            // but might have diagnostic capabilities.
                                            // We default to 1 channel if we don't know the channel count,
                                            // or we try to derive it from the name if possible.
                                            // A better way is to find a channel count property, but if not:
                                            maxChannelCount = 1;
                                        }

                                        string uniqueKey = $"{location}_{attr.Name}";
                                        if (!seenAttrs.Contains(uniqueKey))
                                        {
                                            seenAttrs.Add(uniqueKey);
                                            
                                            // Format for CSV: Escape commas in values
                                            string cleanVal = val.ToString().Replace(",", ";");
                                            lock (csvReportData)
                                            {
                                                csvReportData.Add($"{stationName},{itemName},{typeId},{location},{attr.Name},{cleanVal}");
                                            }

                                            string displayLoc = location == "Module" ? "" : $"[{location}]";
                                            Log($"  [{logPrefix}] {itemName}{displayLoc} -> {attr.Name}: {val}");
                                        }
                                    }
                                }
                                catch { /* Attribute is protected or invalid, skip silently */ }
                            }
                        }
                    }
                }
                catch { }
            };

            // =================================================================
            // METHOD 1: Direct Attributes on the Item
            // Catches ST modules and flattened indexed arrays (e.g. Channel[0].Diag)
            // =================================================================
            checkAndLogAttrs(item, "MOD-DIAG", "Module");

            // =================================================================
            // METHOD 2: Native 'Channels' Property
            // Catches ET200SP High Feature (HF) sub-nodes
            // =================================================================
            bool isHighFeature = false;
            try
            {
                dynamic channels = item.Channels;
                if (channels != null)
                {
                    isHighFeature = true;
                    int idx = 0;
                    // Use exception-based smart probing to count actual channels
                    // since looping through all possible channels causes massive lag
                    for (int i = 0; i < 64; i++) // Arbitrary max
                    {
                        try
                        {
                            dynamic ch = channels[i];
                            checkAndLogAttrs(ch, "CH-PROP-DIAG", $"Ch_{i}");
                            idx++;
                            maxChannelCount = Math.Max(maxChannelCount, idx);
                        }
                        catch
                        {
                            break; // Exception thrown -> no more channels
                        }
                    }
                }
            }
            catch { }

            // =================================================================
            // METHOD 3: ChannelProvider Service
            // Catches integrated IO on newer S7-1200/1500 firmware
            // =================================================================
            try
            {
                dynamic provider = item.GetService("Siemens.Engineering.Hw.Features.ChannelProvider");
                if (provider != null)
                {
                    isHighFeature = true;
                    dynamic svcChannels = provider.Channels;
                    if (svcChannels != null)
                    {
                        int idx = 0;
                        for (int i = 0; i < 64; i++) // Arbitrary max
                        {
                            try
                            {
                                dynamic ch = svcChannels[i];
                                checkAndLogAttrs(ch, "CH-SVC-DIAG", $"Ch_{i}");
                                idx++;
                                maxChannelCount = Math.Max(maxChannelCount, idx);
                            }
                            catch
                            {
                                break; // Exception thrown -> no more channels
                            }
                        }
                    }
                }
            }
            catch { }

            // Add to cache if diagnostics were found
            if (hasDiagnostics)
            {
                // Ensure at least 1 channel if diagnostics found
                if (maxChannelCount == 0) maxChannelCount = 1;

                lock (diagnosticModules)
                {
                    // Avoid duplicates
                    bool exists = false;
                    foreach (var mod in diagnosticModules)
                    {
                        if (mod.ModuleName == itemName && mod.HardwareIdentifier == hwId)
                        {
                            exists = true;
                            // Update max channels if this pass found more
                            mod.ChannelCount = Math.Max(mod.ChannelCount, maxChannelCount);
                            break;
                        }
                    }

                    if (!exists)
                    {
                        diagnosticModules.Add(new DiagnosticModuleInfo
                        {
                            ModuleName = itemName,
                            HardwareIdentifier = hwId,
                            ChannelCount = maxChannelCount,
                            IsHighFeature = isHighFeature
                        });
                        Log($"[CACHE] Cached module '{itemName}' (HW_IO: {hwId}, Channels: {maxChannelCount}, HF: {isHighFeature}) for SCL generation.");
                    }
                }
            }
        }

        private void BtnExportCsv_Click(object? sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog { Filter = "CSV File|*.csv" };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllLines(sfd.FileName, csvReportData);
                Log($"\n[EXPORT] Data saved to {sfd.FileName}");
            }
        }

        private void BtnDisconnect_Click(object? sender, EventArgs e)
        {
            tiaPortal?.Dispose();
            tiaPortal = null;
            btnConnect.Enabled = true;
            btnDisconnect.Enabled = false;
            btnExportCsv.Enabled = false;
            Log("Disconnected from TIA Portal.");
        }

        private void Log(string msg)
        {
            if (this.InvokeRequired) this.Invoke(new Action(() => Log(msg)));
            else
            {
                txtOutput.AppendText(msg + Environment.NewLine);
                txtOutput.ScrollToCaret();
            }
        }
    }
}