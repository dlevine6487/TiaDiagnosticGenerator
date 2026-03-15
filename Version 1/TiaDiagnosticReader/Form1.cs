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
        private RichTextBox txtOutput;
        private List<string> csvReportData;

        public Form1()
        {
            this.Text = "TIA Portal V20 Ultimate Diagnostic Scanner";
            this.Size = new System.Drawing.Size(1000, 800);
            this.csvReportData = new List<string>();

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
            this.Controls.Add(txtOutput);
        }

        private async void BtnConnect_Click(object? sender, EventArgs e)
        {
            txtOutput.Clear();
            csvReportData.Clear();
            csvReportData.Add("Station,Item,Type,Location,Attribute,Value");

            btnConnect.Enabled = false;
            Log("Initializing TIA Portal connection on background thread...");

            await Task.Run(() => PerformScan());

            btnDisconnect.Enabled = (tiaPortal != null);
            btnExportCsv.Enabled = (csvReportData.Count > 1);
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
            try { itemName = item.Name.ToString(); } catch { }
            try { typeId = item.TypeIdentifier.ToString(); } catch { }

            // Keywords we are looking for in the attribute names
            string[] keywords = { "diag", "valuestatus", "wirebreak", "shortcircuit", "overflow", "underflow", "nosupplyvoltage" };
            
            // Track seen attributes to prevent duplicate logging if multiple methods find the same data
            HashSet<string> seenAttrs = new HashSet<string>();

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
            try
            {
                dynamic channels = item.Channels;
                if (channels != null)
                {
                    int idx = 0;
                    foreach (dynamic ch in channels)
                    {
                        checkAndLogAttrs(ch, "CH-PROP-DIAG", $"Ch_{idx}");
                        idx++;
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
                    dynamic svcChannels = provider.Channels;
                    if (svcChannels != null)
                    {
                        int idx = 0;
                        foreach (dynamic ch in svcChannels)
                        {
                            checkAndLogAttrs(ch, "CH-SVC-DIAG", $"Ch_{idx}");
                            idx++;
                        }
                    }
                }
            }
            catch { }
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