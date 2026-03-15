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
        private ComboBox cmbPlcSelect;
        private Label lblPlcSelect;
        private RichTextBox txtOutput;
        private List<string> csvReportData;
        private List<DiagnosticModuleInfo> diagnosticModules;
        private Dictionary<string, string> ioSystemToPlc; // Maps IO System Name -> Controlling PLC Name

        public class DiagnosticModuleInfo
        {
            public string StationName { get; set; } = "";
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
            this.ioSystemToPlc = new Dictionary<string, string>();

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

            lblPlcSelect = new Label();
            lblPlcSelect.Text = "Select PLC:";
            lblPlcSelect.Location = new System.Drawing.Point(410, 13);
            lblPlcSelect.Width = 70;

            cmbPlcSelect = new ComboBox();
            cmbPlcSelect.Location = new System.Drawing.Point(480, 10);
            cmbPlcSelect.Width = 200;
            cmbPlcSelect.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbPlcSelect.Enabled = false;

            btnGenerateScl = new Button();
            btnGenerateScl.Text = "Generate SCL";
            btnGenerateScl.Location = new System.Drawing.Point(690, 10);
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
            this.Controls.Add(lblPlcSelect);
            this.Controls.Add(cmbPlcSelect);
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

            string selectedStation = cmbPlcSelect.SelectedItem?.ToString() ?? "";

            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select a folder to save the generated S7DCL/S7res files";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    btnGenerateScl.Enabled = false;
                    Log($"Generating SCL/S7DCL files for station: {selectedStation}...");

                    await Task.Run(() =>
                    {
                        try
                        {
                            GenerateAndSaveVciFiles(fbd.SelectedPath, selectedStation);
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

        private void GenerateAndSaveVciFiles(string outDir, string selectedStation)
        {
            string udtDir = Path.Combine(outDir, "UDT");
            Directory.CreateDirectory(udtDir);

            // Filter modules for the selected station
            var stationModules = new List<DiagnosticModuleInfo>();
            foreach (var mod in diagnosticModules)
            {
                if (string.IsNullOrEmpty(selectedStation) || mod.StationName == selectedStation)
                {
                    stationModules.Add(mod);
                }
            }

            if (stationModules.Count == 0)
            {
                Log($"No diagnostic modules found for station {selectedStation}.");
                return;
            }

            // Find max channel count
            int maxChannelsAll = 1;
            foreach (var mod in stationModules)
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

            // 6. FB 1500Diag82
            System.Text.StringBuilder fbScl = new System.Text.StringBuilder();
            fbScl.AppendLine("FUNCTION_BLOCK \"1500Diag82\"");
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
            File.WriteAllText(Path.Combine(outDir, "1500Diag82.scl"), fbScl.ToString());

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
            if (stationModules.Count > 0)
            {
                tagsS7dcl.AppendLine($"            diag82 : Array[1..{stationModules.Count}] of \"typeDiag82\";");
            }

            int idx = 1;
            foreach (var mod in stationModules)
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

            idx = 1; // Reuse the previously defined idx variable
            int compileUnitId = 3; // Global ID in the XML tree for CompileUnits
            int multiTextId = 4;   // Global ID in the XML tree for MultilingualText
            int uidBase = 20;      // Global UId for FlgNet elements across all networks

            foreach (var mod in stationModules)
            {
                string safeName = System.Text.RegularExpressions.Regex.Replace(mod.ModuleName, @"[^a-zA-Z0-9_]", "");
                if (string.IsNullOrEmpty(safeName) || char.IsDigit(safeName[0])) safeName = "Mod" + safeName + "_" + idx;

                obXml.AppendLine($"      <SW.Blocks.CompileUnit ID=\"{compileUnitId:X}\" CompositionName=\"CompileUnits\">");
                obXml.AppendLine("        <AttributeList>");
                obXml.AppendLine("          <NetworkSource><FlgNet xmlns=\"http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v5\">");
                obXml.AppendLine("  <Parts>");

                // Hardware constant access using the literal integer value cached from the audit
                obXml.AppendLine($"    <Access Scope=\"LiteralConstant\" UId=\"{uidBase + 1}\">");
                obXml.AppendLine($"      <Constant>");
                obXml.AppendLine($"        <ConstantType>HW_IO</ConstantType>");
                obXml.AppendLine($"        <ConstantValue>{mod.HardwareIdentifier}</ConstantValue>");
                obXml.AppendLine($"      </Constant>");
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
                obXml.AppendLine("      <CallInfo Name=\"1500Diag82\" BlockType=\"FB\">");
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

                obXml.AppendLine("  </Parts>");
                obXml.AppendLine("  <Wires>");

                obXml.AppendLine($"    <Wire UId=\"{uidBase + 5}\">");
                obXml.AppendLine($"      <OpenCon UId=\"{uidBase + 6}\" />");
                obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"en\" />");
                obXml.AppendLine("    </Wire>");
                obXml.AppendLine($"    <Wire UId=\"{uidBase + 7}\">");
                obXml.AppendLine($"      <IdentCon UId=\"{uidBase + 1}\" />");
                obXml.AppendLine($"      <NameCon UId=\"{uidBase + 3}\" Name=\"fId\" />");
                obXml.AppendLine("    </Wire>");

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

                // Increment uidBase significantly to ensure no overlaps in next network
                uidBase += 30 + outVars.Length * 2;

                obXml.AppendLine("  </Wires>");
                obXml.AppendLine("</FlgNet></NetworkSource>");
                obXml.AppendLine("          <ProgrammingLanguage>FBD</ProgrammingLanguage>");
                obXml.AppendLine("        </AttributeList>");
                obXml.AppendLine("        <ObjectList>");
                obXml.AppendLine($"          <MultilingualText ID=\"{multiTextId:X}\" CompositionName=\"Comment\">");
                obXml.AppendLine("            <ObjectList>");
                obXml.AppendLine($"              <MultilingualTextItem ID=\"{(multiTextId+1):X}\" CompositionName=\"Items\">");
                obXml.AppendLine("                <AttributeList>");
                obXml.AppendLine("                  <Culture>en-US</Culture>");
                obXml.AppendLine($"                  <Text>Diagnostics for {safeName}</Text>");
                obXml.AppendLine("                </AttributeList>");
                obXml.AppendLine("              </MultilingualTextItem>");
                obXml.AppendLine("            </ObjectList>");
                obXml.AppendLine("          </MultilingualText>");
                obXml.AppendLine($"          <MultilingualText ID=\"{(multiTextId+2):X}\" CompositionName=\"Title\">");
                obXml.AppendLine("            <ObjectList>");
                obXml.AppendLine($"              <MultilingualTextItem ID=\"{(multiTextId+3):X}\" CompositionName=\"Items\">");
                obXml.AppendLine("                <AttributeList>");
                obXml.AppendLine("                  <Culture>en-US</Culture>");
                obXml.AppendLine($"                  <Text>Network {idx}: {mod.ModuleName}</Text>");
                obXml.AppendLine("                </AttributeList>");
                obXml.AppendLine("              </MultilingualTextItem>");
                obXml.AppendLine("            </ObjectList>");
                obXml.AppendLine("          </MultilingualText>");
                obXml.AppendLine("        </ObjectList>");
                obXml.AppendLine("      </SW.Blocks.CompileUnit>");

                compileUnitId += 5; // Increment to prevent collisions in document
                multiTextId += 5;

                // Create instance DB (s7dcl) for each card
                string instDb = $@"        {{
           S7_Optimized := ""TRUE"";
           S7_StandardRetain := ""FALSE"";
           S7_Version := ""0.1""
        }}
    DATA_BLOCK Inst_{safeName} : ""1500Diag82""
    END_DATA_BLOCK";
                File.WriteAllText(Path.Combine(outDir, $"Inst_{safeName}.s7dcl"), instDb);

                idx++;
            }

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
            cmbPlcSelect.Items.Clear();
            if (ioSystemToPlc == null) ioSystemToPlc = new Dictionary<string, string>();
            ioSystemToPlc.Clear();

            btnConnect.Enabled = false;
            Log("Initializing TIA Portal connection on background thread...");

            await Task.Run(() => PerformScan());

            btnDisconnect.Enabled = (tiaPortal != null);
            btnExportCsv.Enabled = (csvReportData.Count > 1);

            if (diagnosticModules.Count > 0)
            {
                HashSet<string> uniquePlcs = new HashSet<string>();
                foreach (var m in diagnosticModules) uniquePlcs.Add(m.StationName);

                foreach (var plc in uniquePlcs) cmbPlcSelect.Items.Add(plc);

                if (cmbPlcSelect.Items.Count > 0) cmbPlcSelect.SelectedIndex = 0;

                cmbPlcSelect.Enabled = true;
                btnGenerateScl.Enabled = true;
            }

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

                Log("\n--- PASS 1: Mapping IO Systems to PLCs ---");
                foreach (dynamic device in project.Devices) BuildIoSystemMap(device);
                BuildIoSystemMapForGroups(project.DeviceGroups);
                if (project.UngroupedDevicesGroup != null)
                {
                    foreach (dynamic device in project.UngroupedDevicesGroup.Devices) BuildIoSystemMap(device);
                }

                Log("\n--- PASS 2: Hardware Audit ---");
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

        private string GetPlcName(dynamic deviceItems, string defaultName)
        {
            if (deviceItems == null) return defaultName;
            foreach (dynamic item in deviceItems)
            {
                if (item != null)
                {
                    bool isPlc = false;
                    try
                    {
                        string classification = item.Classification.ToString();
                        if (classification.IndexOf("CPU", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            isPlc = true;
                        }
                    }
                    catch { }

                    if (!isPlc)
                    {
                        try
                        {
                            var swContainer = item.GetService("Siemens.Engineering.SW.SoftwareContainer");
                            if (swContainer != null)
                            {
                                isPlc = true;
                            }
                        }
                        catch { }
                    }

                    if (!isPlc)
                    {
                        try
                        {
                            var plcContainer = item.GetService("Siemens.Engineering.SW.PlcSoftware");
                            if (plcContainer != null)
                            {
                                isPlc = true;
                            }
                        }
                        catch { }
                    }

                    if (isPlc)
                    {
                        try { return item.Name.ToString(); } catch { }
                    }

                    try
                    {
                        if (item.DeviceItems != null)
                        {
                            string subResult = GetPlcName(item.DeviceItems, defaultName);
                            if (subResult != defaultName) return subResult;
                        }
                    }
                    catch { }
                }
            }
            return defaultName;
        }

        private void BuildIoSystemMap(dynamic device)
        {
            if (device == null) return;
            string deviceName = "Unknown";
            try { deviceName = device.Name.ToString(); } catch { }

            // Try to find the actual PLC name within the station
            string plcName = GetPlcName(device.DeviceItems, deviceName);

            // Note: deviceName is the top level container, plcName is the specific CPU inside it.
            // When mapping IO systems, we want them mapped to the specific PLC name.
            MapIoSystemsRecursive(device.DeviceItems, plcName);
        }

        private void BuildIoSystemMapForGroups(dynamic groups)
        {
            if (groups == null) return;
            foreach (dynamic group in groups)
            {
                foreach (dynamic device in group.Devices) BuildIoSystemMap(device);
                if (group.Groups != null) BuildIoSystemMapForGroups(group.Groups);
            }
        }

        private void MapIoSystemsRecursive(dynamic items, string deviceName)
        {
            if (items == null) return;
            foreach (dynamic item in items)
            {
                if (item != null)
                {
                    try
                    {
                        var networkInterface = item.GetService("Siemens.Engineering.HW.Features.NetworkInterface");
                        if (networkInterface != null)
                        {
                            var ioControllers = networkInterface.IoControllers;
                            if (ioControllers != null)
                            {
                                foreach (var controller in ioControllers)
                                {
                                    var ioSystem = controller.IoSystem;
                                    if (ioSystem != null)
                                    {
                                        string ioSystemName = ioSystem.Name.ToString();
                                        if (!ioSystemToPlc.ContainsKey(ioSystemName))
                                        {
                                            ioSystemToPlc[ioSystemName] = deviceName;
                                            Log($"[MAP] IO System '{ioSystemName}' -> PLC '{deviceName}'");
                                        }
                                    }
                                }
                            }

                            // Handle Profibus DP Masters as well, if applicable
                            try
                            {
                                var dpMasters = networkInterface.DpMasters;
                                if (dpMasters != null)
                                {
                                    foreach (var master in dpMasters)
                                    {
                                        var dpSystem = master.DpSystem;
                                        if (dpSystem != null)
                                        {
                                            string dpSystemName = dpSystem.Name.ToString();
                                            if (!ioSystemToPlc.ContainsKey(dpSystemName))
                                            {
                                                ioSystemToPlc[dpSystemName] = deviceName;
                                                Log($"[MAP] DP System '{dpSystemName}' -> PLC '{deviceName}'");
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

                    if (item.DeviceItems != null)
                    {
                        MapIoSystemsRecursive(item.DeviceItems, deviceName);
                    }
                }
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

            // Try to find the actual PLC name within the station first
            string plcName = GetPlcName(device.DeviceItems, stationName);

            // Override stationName if this device is part of a mapped IO System
            string controllingPlcName = ResolveControllingPlc(device.DeviceItems, plcName);

            // Important: we use the controllingPlcName (which defaults to the actual PLC name if not remote IO)
            Log($"\n>>> STATION: {controllingPlcName}");
            RecursiveWalk(device.DeviceItems, controllingPlcName);
        }

        private string ResolveControllingPlc(dynamic items, string defaultName)
        {
            if (items == null) return defaultName;

            foreach (dynamic item in items)
            {
                if (item != null)
                {
                    try
                    {
                        var networkInterface = item.GetService("Siemens.Engineering.HW.Features.NetworkInterface");
                        if (networkInterface != null)
                        {
                            var ioConnectors = networkInterface.IoConnectors;
                            if (ioConnectors != null)
                            {
                                foreach (var connector in ioConnectors)
                                {
                                    var ioSystem = connector.ConnectedToIoSystem;
                                    if (ioSystem != null)
                                    {
                                        string ioSystemName = ioSystem.Name.ToString();
                                        if (ioSystemToPlc.ContainsKey(ioSystemName))
                                        {
                                            return ioSystemToPlc[ioSystemName]; // Return the controlling PLC
                                        }
                                    }
                                }
                            }

                            try
                            {
                                var dpSlaves = networkInterface.DpSlaves;
                                if (dpSlaves != null)
                                {
                                    foreach (var slave in dpSlaves)
                                    {
                                        var dpSystem = slave.ConnectedToDpSystem;
                                        if (dpSystem != null)
                                        {
                                            string dpSystemName = dpSystem.Name.ToString();
                                            if (ioSystemToPlc.ContainsKey(dpSystemName))
                                            {
                                                return ioSystemToPlc[dpSystemName];
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

                    if (item.DeviceItems != null)
                    {
                        string subResult = ResolveControllingPlc(item.DeviceItems, defaultName);
                        if (subResult != defaultName) return subResult; // Found a map match deeper down
                    }
                }
            }

            return defaultName;
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

            // Filter out non-IO components from generating instances
            string itemNameLower = itemName.ToLower();
            if (itemNameLower.Contains("server module") ||
                itemNameLower.Contains("profinet interface") ||
                itemNameLower.Contains("card reader/writer") ||
                itemNameLower.Contains("cpu "))
            {
                return; // Exclude
            }

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
                            StationName = stationName,
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