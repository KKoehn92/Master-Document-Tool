using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MDL_TEST_MA
{
    public partial class EingabeForm : Form
    {
        private Panel scrollPanel;
        private TableLayoutPanel tlp;
        private ProgressBar progressBar;
        private ToolTip toolTip;

        private const int OWNED_START_ROW = 7;
        private const string OWNED_LAST_COL = "AF";

        private readonly HashSet<int> checkboxRows = new HashSet<int>
        {
            18,20,22,
            58,60,62,64,66,68,70,72,74,
            80,82,83,85,87,89,93,
            97,99,101,103,107,109,112,115,119,121,
            123,125,127,129,133,137,139
        };

        public EingabeForm()
        {
            InitializeComponent();
            Width = 1000;
            Height = 900;
            Text = "MDL Eingabemaske";

            toolTip = new ToolTip
            {
                AutoPopDelay = 5000,
                InitialDelay = 500,
                ReshowDelay = 500,
                ShowAlways = true
            };

            scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.LightSteelBlue,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(scrollPanel);

            var commonFont = new Font("Segoe UI", 9, FontStyle.Bold);
            var commonPadding = new Padding(12, 6, 12, 6);

            var btnRefresh = new Button
            {
                Text = "Fragen neu laden",
                AutoSize = true,
                Font = commonFont,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Navy,
                ForeColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Padding = commonPadding
            };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (_, __) => BuildInputMask();

            var btnReport = new Button
            {
                Text = "Reporting aktualisieren",
                AutoSize = true,
                Font = commonFont,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Navy,
                ForeColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Padding = commonPadding
            };
            btnReport.FlatAppearance.BorderSize = 0;
            btnReport.Click += (_, __) =>
            {
                try
                {
                    Globals.ThisAddIn.BuildReporting();
                    MessageBox.Show("Reporting aktualisiert.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Aktualisieren des Reportings:\n" + ex.Message,
                                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            Size reportSize = btnReport.GetPreferredSize(Size.Empty);
            btnReport.AutoSize = false; btnReport.Size = reportSize;
            btnRefresh.AutoSize = false; btnRefresh.Size = reportSize;

            btnRefresh.Location = new Point(10, 10);
            btnReport.Location = new Point(btnRefresh.Right + 12, 10);

            scrollPanel.Controls.Add(btnRefresh);
            scrollPanel.Controls.Add(btnReport);

            BuildInputMask();
        }

        private void BuildInputMask()
        {
            var wb = GetActiveWorkbookOrWarn("Eingabemaske aufbauen");
            if (wb == null) return;

            var ws = FindSheet(wb, "Eingabemaske");
            if (ws == null)
            {
                MessageBox.Show(
                    "Das Tabellenblatt „Eingabemaske“ wurde in der aktiven Arbeitsmappe nicht gefunden.\n" +
                    "Bitte öffne die richtige Datei oder benenne das Blatt entsprechend.",
                    "Blatt fehlt", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            scrollPanel.SuspendLayout();
            try
            {
                foreach (var c in scrollPanel.Controls.Cast<Control>().Where(c => !(c is Button)).ToList())
                    scrollPanel.Controls.Remove(c);

                int topOffset = 50;
                var btns = scrollPanel.Controls.OfType<Button>().ToList();
                if (btns.Any())
                    topOffset = btns.Max(b => b.Bottom) + 20;

                tlp = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Location = new Point(10, topOffset),
                    CellBorderStyle = TableLayoutPanelCellBorderStyle.None
                };
                tlp.SuspendLayout();
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
                scrollPanel.Controls.Add(tlp);

                int row = 0;
                for (int z = 4; z <= 139; z++)
                {
                    string frage = Convert.ToString(ws.Cells[z, 1].Value2);
                    string typ = Convert.ToString(ws.Cells[z, 4].Value2);
                    if (string.IsNullOrWhiteSpace(frage)) continue;

                    tlp.RowCount++;
                    tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                    var lbl = new Label
                    {
                        Text = frage,
                        AutoSize = true,
                        MaximumSize = new Size(650, 0),
                        Font = new Font("Segoe UI", 10, FontStyle.Bold),
                        Anchor = AnchorStyles.Left | AnchorStyles.Top,
                        Tag = z
                    };
                    tlp.Controls.Add(lbl, 0, row);

                    if (z == 30)
                    {
                        tlp.Controls.Add(BuildZeile30(z), 1, row);
                        row++; continue;
                    }
                    if (frage.Equals("Gibt es Provisorien", StringComparison.OrdinalIgnoreCase))
                    {
                        BuildProvisorien(z, ref row);
                        continue;
                    }
                    if (z == 80)
                    {
                        tlp.Controls.Add(BuildCheckBox(z), 1, row);
                        row++; continue;
                    }
                    if (z == 81 || z == 105 || z == 131 || z == 135)
                    {
                        tlp.Controls.Add(BuildTextBox(z), 1, row);
                        row++; continue;
                    }

                    Control ctrl;
                    if (z == 20)
                        ctrl = BuildTextBox(z);
                    else if (z == 26 || checkboxRows.Contains(z)
                             || (!string.IsNullOrEmpty(typ) && typ.Equals("ja oder nein", StringComparison.OrdinalIgnoreCase)))
                        ctrl = BuildCheckBox(z);
                    else if ((!string.IsNullOrEmpty(typ) && typ.Equals("zahlenwert", StringComparison.OrdinalIgnoreCase)) || z == 28)
                        ctrl = BuildNumeric(z);
                    else
                        ctrl = BuildTextBox(z);

                    tlp.Controls.Add(ctrl, 1, row);
                    row++;
                }

                var wsLbl = FindSheet(wb, "Eingabemaske");
                if (wsLbl != null)
                {
                    foreach (var l in tlp.Controls.OfType<Label>())
                        if (l.Tag is int t)
                            wsLbl.Cells[t, 1].Value2 = l.Text;
                }

                int last = tlp.RowCount;
                tlp.RowCount++;
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var btnSave = new Button
                {
                    Text = "Speichern",
                    Width = 120,
                    Height = 35,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.Navy,
                    ForeColor = Color.White,
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };
                btnSave.FlatAppearance.BorderSize = 0;
                btnSave.Click += SpeichernButton_Click;
                tlp.Controls.Add(btnSave, 0, last);

                progressBar = new ProgressBar
                {
                    Minimum = 0,
                    Maximum = 100,
                    Value = 0,
                    Width = 300,
                    Height = 20,
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };
                tlp.Controls.Add(progressBar, 1, last);

                tlp.ResumeLayout();
            }
            finally
            {
                scrollPanel.ResumeLayout();
            }
        }

        private void SpeichernButton_Click(object sender, EventArgs e)
        {
            var wb = GetActiveWorkbookOrWarn("Speichern");
            if (wb == null) return;

            var wsOut = FindSheet(wb, "Master Document List");
            if (wsOut == null)
            {
                MessageBox.Show("Blatt „Master Document List“ nicht gefunden.", "Blatt fehlt",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var wsOhl = FindSheet(wb, "IBL-OHL Neubau") ?? FindSheet(wb, "IHL-OHL Neubau");
            if (wsOhl == null)
            {
                MessageBox.Show("Blatt „IBL-/IHL-OHL Neubau“ nicht gefunden.", "Blatt fehlt",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var old = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                var ctrls = GetAllControls(tlp).Where(c => c.Tag != null).ToList();
                var map = ctrls.GroupBy(c => c.Tag).ToDictionary(g => g.Key, g => g.Last());

                string RV(int t)
                {
                    if (!map.TryGetValue(t, out Control c)) return "";
                    if (c is TextBox tb) return tb.Text;
                    if (c is NumericUpDown nu) return nu.Value.ToString();
                    if (c is CheckBox cb) return cb.Checked ? "Ja" : "Nein";
                    return "";
                }

                string A4 = RV(4), A5 = RV(5), A6 = RV(6), A8 = RV(8),
                       A10 = RV(10), A12 = RV(12), A14 = RV(14), A16 = RV(16),
                       A20 = RV(20), A22 = RV(22), A24 = RV(24), A26 = RV(26),
                       A28 = RV(28), A58 = RV(58), A60 = RV(60), A62 = RV(62),
                       A64 = RV(64), A66 = RV(66), A68 = RV(68), A70 = RV(70),
                       A72 = RV(72), A74 = RV(74), A80 = RV(80), A81 = RV(81),
                       A82 = RV(82), A83 = RV(83), A85 = RV(85), A87 = RV(87),
                       A89 = RV(89), A93 = RV(93), A97 = RV(97), A99 = RV(99),
                       A101 = RV(101), A103 = RV(103), A105 = RV(105), A107 = RV(107),
                       A109 = RV(109), A112 = RV(112), A115 = RV(115), A119 = RV(119),
                       A121 = RV(121), A123 = RV(123), A125 = RV(125), A127 = RV(127),
                       A129 = RV(129), A131 = RV(131), A133 = RV(133), A135 = RV(135),
                       A137 = RV(137), A139 = RV(139);

                if (progressBar != null) progressBar.Value = 25;

                if (A133.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                    wsOut.Cells[3, 6].Value2 = A135;
                else
                    wsOut.Cells[3, 6].Value2 = "Nein";

                if (!string.IsNullOrWhiteSpace(A10))
                    wsOut.Cells[4, 4].Value2 = A10;
                if (!string.IsNullOrWhiteSpace(A20))
                    wsOut.Cells[1, 4].Value2 = A20;

                wsOut.Cells[3, 4].Value2 =
                    A22.Equals("Ja", StringComparison.OrdinalIgnoreCase) ? A24 : "keine Mitnahme";
                wsOut.Cells[2, 4].Value2 =
                    A26.Equals("Ja", StringComparison.OrdinalIgnoreCase) ? A28 : "keine Provisorien";

                var mastList = ExpandMastTokens(A14);

                bool IsJe(string v)
                {
                    if (string.IsNullOrWhiteSpace(v)) return false;
                    v = v.Trim();
                    return v.Equals("Je Mast", StringComparison.OrdinalIgnoreCase)
                        || v.StartsWith("Je Mast", StringComparison.OrdinalIgnoreCase)
                        || v.StartsWith("Je Masttyp", StringComparison.OrdinalIgnoreCase);
                }

                var rows = new List<Dictionary<int, string>>();

                void AddStdBlock(int srcRow)
                {
                    string rawE = wsOhl.Cells[srcRow, 5].Value2?.ToString() ?? "";
                    string rawM = wsOhl.Cells[srcRow, 13].Value2?.ToString() ?? "";
                    string rawJ = wsOhl.Cells[srcRow, 10].Value2?.ToString() ?? "";
                    string rawW = wsOhl.Cells[srcRow, 23].Value2?.ToString() ?? "";
                    string rawP = wsOhl.Cells[srcRow, 16].Value2?.ToString() ?? "";
                    string rawV = wsOhl.Cells[srcRow, 22].Value2?.ToString() ?? "";
                    string rawC = wsOhl.Cells[srcRow, 3].Value2?.ToString() ?? "";

                    bool fillK = IsJe(rawV);
                    var listK = fillK ? mastList : new List<string> { "" };

                    foreach (string m in listK)
                        rows.Add(new Dictionary<int, string>
                        {
                            [ExcelSpalteZuIndex("A")] = rawE,
                            [ExcelSpalteZuIndex("B")] = rawM,
                            [ExcelSpalteZuIndex("F")] = A8,
                            [ExcelSpalteZuIndex("H")] = rawE + "_" + A16,
                            [ExcelSpalteZuIndex("I")] = A4,
                            [ExcelSpalteZuIndex("J")] = A16,
                            [ExcelSpalteZuIndex("K")] = fillK ? m : "",
                            [ExcelSpalteZuIndex("L")] = "Ungeprüft",
                            [ExcelSpalteZuIndex("M")] = "Zur Prüfung und Freigabe",
                            [ExcelSpalteZuIndex("T")] = A8,
                            [ExcelSpalteZuIndex("U")] = A5,
                            [ExcelSpalteZuIndex("V")] = A4,
                            [ExcelSpalteZuIndex("Y")] = "Deutsch",
                            [ExcelSpalteZuIndex("AA")] = "C2",
                            [ExcelSpalteZuIndex("AB")] = rawJ,
                            [ExcelSpalteZuIndex("AC")] = rawW,
                            [ExcelSpalteZuIndex("AD")] = rawP,
                            [ExcelSpalteZuIndex("AE")] = rawV,
                            [ExcelSpalteZuIndex("AF")] = rawC
                        });
                }

                if (A58.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(7);
                if (A60.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(11);
                if (A62.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                {
                    var e14 = wsOhl.Cells[14, 5].Value2?.ToString() ?? "";
                    var e30 = wsOhl.Cells[30, 5].Value2?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(e14) && !e14.Equals(e30, StringComparison.OrdinalIgnoreCase))
                        AddStdBlock(14);
                }
                if (A64.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(24);
                if (A66.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(33);
                if (A68.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(36);
                if (A70.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(41);
                if (A72.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(42);
                if (A74.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                {
                    AddStdBlock(44); AddStdBlock(46); AddStdBlock(48);
                    AddStdBlock(49); AddStdBlock(76); AddStdBlock(45);
                }

                if (A80.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                {
                    var list81 = ExpandMastTokens(A81);
                    AddSources(new[] { 58 }, list81);
                }

                if (A82.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(54);
                if (A83.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(55); AddStdBlock(56); }
                if (A85.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(59);
                if (A87.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(62); AddStdBlock(63); }
                if (A89.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(64); AddStdBlock(87); AddStdBlock(83); AddStdBlock(92); }
                if (A93.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(69); AddStdBlock(127); }
                if (A97.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(70);
                if (A99.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(72); AddStdBlock(43); }
                if (A101.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(73);

                if (A103.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                {
                    var list105 = ExpandMastTokens(A105);
                    AddSources(new[] { 77 }, list105);
                }

                if (A107.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(100); AddStdBlock(121); }
                if (A109.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(102); AddStdBlock(131); }
                if (A112.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(103); AddStdBlock(104); }
                if (A115.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(112);
                if (A119.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(125);
                if (A121.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(126);
                if (A123.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(128);
                if (A125.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(130);
                if (A127.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(133);
                if (A137.Equals("Ja", StringComparison.OrdinalIgnoreCase)) { AddStdBlock(139); AddStdBlock(140); }
                if (A139.Equals("Ja", StringComparison.OrdinalIgnoreCase)) AddStdBlock(142);

                if (!string.IsNullOrWhiteSpace(A6))
                {
                    int[] srcsA6 = { 10, 37, 38, 39, 47, 50, 52, 80, 81, 82, 95, 96, 97, 98, 99, 120 };
                    foreach (int src in srcsA6)
                    {
                        string rawE = wsOhl.Cells[src, 5].Value2?.ToString() ?? "";
                        string rawM = wsOhl.Cells[src, 13].Value2?.ToString() ?? "";
                        string rawJ = wsOhl.Cells[src, 10].Value2?.ToString() ?? "";
                        string rawW = wsOhl.Cells[src, 23].Value2?.ToString() ?? "";
                        string rawP = wsOhl.Cells[src, 16].Value2?.ToString() ?? "";
                        string rawV = wsOhl.Cells[src, 22].Value2?.ToString() ?? "";
                        string rawC = wsOhl.Cells[src, 3].Value2?.ToString() ?? "";

                        bool je = IsJe(rawV);
                        var list = je ? mastList : new List<string> { "" };

                        foreach (string m in list)
                            rows.Add(new Dictionary<int, string>
                            {
                                [ExcelSpalteZuIndex("A")] = rawE,
                                [ExcelSpalteZuIndex("B")] = rawM,
                                [ExcelSpalteZuIndex("AB")] = rawJ,
                                [ExcelSpalteZuIndex("AC")] = rawW,
                                [ExcelSpalteZuIndex("AD")] = rawP,
                                [ExcelSpalteZuIndex("AE")] = rawV,
                                [ExcelSpalteZuIndex("AF")] = rawC,
                                [ExcelSpalteZuIndex("AA")] = "C2",
                                [ExcelSpalteZuIndex("Y")] = "Deutsch",
                                [ExcelSpalteZuIndex("M")] = "Zur Prüfung und Freigabe",
                                [ExcelSpalteZuIndex("L")] = "Ungeprüft",
                                [ExcelSpalteZuIndex("U")] = A5,
                                [ExcelSpalteZuIndex("I")] = A4,
                                [ExcelSpalteZuIndex("J")] = A16,
                                [ExcelSpalteZuIndex("V")] = A4,
                                [ExcelSpalteZuIndex("F")] = A6,
                                [ExcelSpalteZuIndex("T")] = A6,
                                [ExcelSpalteZuIndex("H")] = rawE + "_" + A16,
                                [ExcelSpalteZuIndex("K")] = m
                            });
                    }
                }

                if (!string.IsNullOrWhiteSpace(A8))
                {
                    int[] srcsA8 = {
                        3,4,5,6,8,9,12,13,18,26,32,40,57,60,61,65,
                        66,67,68,71,74,75,78,79,84,85,86,88,90,
                        93,94,101,105,106,107,108,109,110,111,
                        113,114,115,116,117,118,119,122,123,124,
                        129,133,134,135,136,137
                    };
                    foreach (int src in srcsA8)
                    {
                        string rawE = wsOhl.Cells[src, 5].Value2?.ToString() ?? "";
                        string rawM = wsOhl.Cells[src, 13].Value2?.ToString() ?? "";
                        string rawJ = wsOhl.Cells[src, 10].Value2?.ToString() ?? "";
                        string rawW = wsOhl.Cells[src, 23].Value2?.ToString() ?? "";
                        string rawP = wsOhl.Cells[src, 16].Value2?.ToString() ?? "";
                        string rawV = wsOhl.Cells[src, 22].Value2?.ToString() ?? "";
                        string rawC = wsOhl.Cells[src, 3].Value2?.ToString() ?? "";

                        bool je = IsJe(rawV);
                        var list = je ? mastList : new List<string> { "" };

                        foreach (string m in list)
                            rows.Add(new Dictionary<int, string>
                            {
                                [ExcelSpalteZuIndex("A")] = rawE,
                                [ExcelSpalteZuIndex("B")] = rawM,
                                [ExcelSpalteZuIndex("AB")] = rawJ,
                                [ExcelSpalteZuIndex("AC")] = rawW,
                                [ExcelSpalteZuIndex("AD")] = rawP,
                                [ExcelSpalteZuIndex("AE")] = rawV,
                                [ExcelSpalteZuIndex("AF")] = rawC,
                                [ExcelSpalteZuIndex("AA")] = "C2",
                                [ExcelSpalteZuIndex("Y")] = "Deutsch",
                                [ExcelSpalteZuIndex("M")] = "Zur Prüfung und Freigabe",
                                [ExcelSpalteZuIndex("L")] = "Ungeprüft",
                                [ExcelSpalteZuIndex("U")] = A5,
                                [ExcelSpalteZuIndex("I")] = A4,
                                [ExcelSpalteZuIndex("J")] = A16,
                                [ExcelSpalteZuIndex("V")] = A4,
                                [ExcelSpalteZuIndex("F")] = A8,
                                [ExcelSpalteZuIndex("T")] = A8,
                                [ExcelSpalteZuIndex("H")] = rawE + "_" + A16,
                                [ExcelSpalteZuIndex("K")] = m
                            });
                    }
                }

                void DoGründung(string tagString, int[] sourceRows)
                {
                    var tb = ctrls.OfType<TextBox>().FirstOrDefault(c => c.Tag is string s && s == tagString);
                    var list = tb != null ? ExpandMastTokens(tb.Text) : new List<string>();

                    foreach (int src in sourceRows)
                    {
                        string rawE = wsOhl.Cells[src, 5].Value2?.ToString() ?? "";
                        string rawM = wsOhl.Cells[src, 13].Value2?.ToString() ?? "";
                        string rawJ = wsOhl.Cells[src, 10].Value2?.ToString() ?? "";
                        string rawW = wsOhl.Cells[src, 23].Value2?.ToString() ?? "";
                        string rawP = wsOhl.Cells[src, 16].Value2?.ToString() ?? "";
                        string rawV = wsOhl.Cells[src, 22].Value2?.ToString() ?? "";
                        string rawC = wsOhl.Cells[src, 3].Value2?.ToString() ?? "";

                        bool fillK = IsJe(rawV);
                        var mastNums = fillK ? list : new List<string> { "" };

                        foreach (string m in mastNums)
                            rows.Add(new Dictionary<int, string>
                            {
                                [ExcelSpalteZuIndex("A")] = rawE,
                                [ExcelSpalteZuIndex("B")] = rawM,
                                [ExcelSpalteZuIndex("AB")] = rawJ,
                                [ExcelSpalteZuIndex("AC")] = rawW,
                                [ExcelSpalteZuIndex("AD")] = rawP,
                                [ExcelSpalteZuIndex("AE")] = rawV,
                                [ExcelSpalteZuIndex("AF")] = rawC,
                                [ExcelSpalteZuIndex("AA")] = "C2",
                                [ExcelSpalteZuIndex("Y")] = "Deutsch",
                                [ExcelSpalteZuIndex("M")] = "Zur Prüfung und Freigabe",
                                [ExcelSpalteZuIndex("L")] = "Ungeprüft",
                                [ExcelSpalteZuIndex("U")] = A5,
                                [ExcelSpalteZuIndex("I")] = A4,
                                [ExcelSpalteZuIndex("J")] = A16,
                                [ExcelSpalteZuIndex("V")] = A4,
                                [ExcelSpalteZuIndex("F")] = A8,
                                [ExcelSpalteZuIndex("T")] = A8,
                                [ExcelSpalteZuIndex("H")] = rawE + "_" + A16,
                                [ExcelSpalteZuIndex("K")] = m
                            });
                    }
                }

                if (ctrls.OfType<CheckBox>().Any(cb => cb.Tag is int t && t == 30 && cb.Text == "Rammgründung" && cb.Checked))
                    DoGründung("TbMast", new[] { 15, 17, 19, 20, 22, 25, 27, 31, 34, 35 });
                if (ctrls.OfType<CheckBox>().Any(cb => cb.Tag is int t && t == 30 && cb.Text == "Bohrgründung" && cb.Checked))
                    DoGründung("TbBohr", new[] { 15, 17, 19, 20, 21, 23, 25, 27, 28 });
                if (ctrls.OfType<CheckBox>().Any(cb => cb.Tag is int t && t == 30 && cb.Text == "Plattengründung" && cb.Checked))
                    DoGründung("TbPlatt", new[] { 15, 17, 19, 20, 27, 29, 30 });

                void AddSources(int[] sourceRows, List<string> mastNumbers)
                {
                    foreach (int src in sourceRows)
                    {
                        string rawE = wsOhl.Cells[src, 5].Value2?.ToString() ?? "";
                        string rawM = wsOhl.Cells[src, 13].Value2?.ToString() ?? "";
                        string rawJ = wsOhl.Cells[src, 10].Value2?.ToString() ?? "";
                        string rawW = wsOhl.Cells[src, 23].Value2?.ToString() ?? "";
                        string rawP = wsOhl.Cells[src, 16].Value2?.ToString() ?? "";
                        string rawV = wsOhl.Cells[src, 22].Value2?.ToString() ?? "";
                        string rawC = wsOhl.Cells[src, 3].Value2?.ToString() ?? "";

                        bool fillK = IsJe(rawV);
                        var listK = fillK ? mastNumbers : new List<string> { "" };

                        foreach (string m in listK)
                            rows.Add(new Dictionary<int, string>
                            {
                                [ExcelSpalteZuIndex("A")] = rawE,
                                [ExcelSpalteZuIndex("B")] = rawM,
                                [ExcelSpalteZuIndex("F")] = A8,
                                [ExcelSpalteZuIndex("H")] = rawE + "_" + A16,
                                [ExcelSpalteZuIndex("I")] = A4,
                                [ExcelSpalteZuIndex("J")] = A16,
                                [ExcelSpalteZuIndex("K")] = m,
                                [ExcelSpalteZuIndex("L")] = "Ungeprüft",
                                [ExcelSpalteZuIndex("M")] = "Zur Prüfung und Freigabe",
                                [ExcelSpalteZuIndex("T")] = A8,
                                [ExcelSpalteZuIndex("U")] = A5,
                                [ExcelSpalteZuIndex("V")] = A4,
                                [ExcelSpalteZuIndex("Y")] = "Deutsch",
                                [ExcelSpalteZuIndex("AA")] = "C2",
                                [ExcelSpalteZuIndex("AB")] = rawJ,
                                [ExcelSpalteZuIndex("AC")] = rawW,
                                [ExcelSpalteZuIndex("AD")] = rawP,
                                [ExcelSpalteZuIndex("AE")] = rawV,
                                [ExcelSpalteZuIndex("AF")] = rawC
                            });
                    }
                }

                if (progressBar != null) progressBar.Value = 60;

                string SV(int r, int col) => Convert.ToString(wsOut.Cells[r, col].Value2);

                int colA = ExcelSpalteZuIndex("A");
                int colH = ExcelSpalteZuIndex("H");
                int colK = ExcelSpalteZuIndex("K");
                int colAF = ExcelSpalteZuIndex(OWNED_LAST_COL);

                var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var used = wsOut.UsedRange;
                int usedStart = used.Row;
                int usedCount = used.Rows.Count;
                int usedEnd = usedStart + usedCount - 1;
                if (usedEnd < OWNED_START_ROW) usedEnd = OWNED_START_ROW - 1;

                for (int r = OWNED_START_ROW; r <= usedEnd; r++)
                {
                    string hv = SV(r, colH);
                    string kv = SV(r, colK);
                    if (!string.IsNullOrWhiteSpace(hv) || !string.IsNullOrWhiteSpace(kv))
                    {
                        existing.Add((hv ?? "") + "|" + (kv ?? ""));
                    }
                }

                var filtered = new List<Dictionary<int, string>>();
                foreach (var d in rows)
                {
                    string hv = d.ContainsKey(colH) ? d[colH] : "";
                    string kv = d.ContainsKey(colK) ? d[colK] : "";
                    string key = hv + "|" + kv;
                    if (!existing.Contains(key))
                    {
                        filtered.Add(d);
                        existing.Add(key);
                    }
                }

                if (filtered.Count == 0)
                {
                    if (progressBar != null) progressBar.Value = 100;
                    MessageBox.Show("Keine neuen Datensätze (alles bereits vorhanden).", "Hinweis",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int lastOwned = OWNED_START_ROW - 1;
                for (int r = usedEnd; r >= OWNED_START_ROW; r--)
                {
                    string a = SV(r, colA);
                    string h = SV(r, colH);
                    if (!string.IsNullOrWhiteSpace(a) || !string.IsNullOrWhiteSpace(h))
                    {
                        lastOwned = r;
                        break;
                    }
                }
                int appendRow = lastOwned + 1;

                int startRowOfAppend = appendRow;
                foreach (var dict in filtered)
                {
                    foreach (var kv in dict)
                    {
                        int c = kv.Key;
                        if (c >= colA && c <= colAF)
                            wsOut.Cells[appendRow, c].Value2 = kv.Value;
                    }
                    appendRow++;
                }

                int firstManualCol = colAF + 1;
                int lastUsedCol = used.Columns.Count + used.Column - 1;
                if (lastUsedCol >= firstManualCol)
                {
                    int lastNewRow = appendRow - 1;
                    wsOut.Range[
                        wsOut.Cells[startRowOfAppend, firstManualCol],
                        wsOut.Cells[lastNewRow, lastUsedCol]
                    ].ClearContents();
                }

                if (progressBar != null) progressBar.Value = 100;

                try { Globals.ThisAddIn.BuildReporting(); } catch { }

                MessageBox.Show($"Daten übernommen! (neu: {filtered.Count})", "Fertig",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                Cursor.Current = old;
            }
        }

        private Control BuildTextBox(int tag) => new TextBox
        {
            Width = 200,
            Font = new Font("Segoe UI", 10),
            Anchor = AnchorStyles.Left | AnchorStyles.Top,
            Tag = tag
        };

        private Control BuildNumeric(int tag) => new NumericUpDown
        {
            Minimum = 0,
            Maximum = 1_000_000,
            Font = new Font("Segoe UI", 10),
            Anchor = AnchorStyles.Left | AnchorStyles.Top,
            Tag = tag
        };

        private Control BuildCheckBox(int tag)
        {
            var cb = new CheckBox
            {
                Font = new Font("Segoe UI", 10),
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Tag = tag
            };
            toolTip.SetToolTip(cb, "Markieren = Ja; leer lassen = Nein");
            return cb;
        }

        private Control BuildZeile30(int tag)
        {
            var pnl = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top
            };

            var cbAll = new CheckBox
            {
                Text = "Alle auswählen",
                AutoSize = true,
                Font = new Font("Segoe UI", 10),
                Tag = tag
            };
            pnl.Controls.Add(cbAll);

            void AddOption(string txt, string tbTag)
            {
                var cb = new CheckBox
                {
                    Text = txt,
                    AutoSize = true,
                    Font = new Font("Segoe UI", 10),
                    Tag = tag
                };
                var tb = new TextBox
                {
                    Width = 150,
                    Font = new Font("Segoe UI", 10),
                    Visible = false,
                    Tag = tbTag
                };
                toolTip.SetToolTip(cb, "Markieren → aktiviert Freitext");
                toolTip.SetToolTip(tb, "Mastnummern, z. B. M001, M003-M010; M015");
                cb.CheckedChanged += (_, __) => tb.Visible = cb.Checked;

                pnl.Controls.Add(cb);
                pnl.Controls.Add(tb);
            }

            AddOption("Rammgründung", "TbMast");
            AddOption("Bohrgründung", "TbBohr");
            AddOption("Plattengründung", "TbPlatt");

            cbAll.CheckedChanged += (_, __) =>
            {
                foreach (var c in pnl.Controls.OfType<CheckBox>())
                {
                    if (!object.ReferenceEquals(c, cbAll))
                        c.Checked = cbAll.Checked;
                }
            };

            return pnl;
        }

        private void BuildProvisorien(int tag, ref int row)
        {
            var chk = (CheckBox)BuildCheckBox(tag);
            tlp.Controls.Add(chk, 1, row); row++;
            tlp.RowCount++; tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var lbl = new Label
            {
                Text = "Anzahl Provisorien:",
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Visible = false
            };
            tlp.Controls.Add(lbl, 0, row);
            var num = (NumericUpDown)BuildNumeric(tag + 1); num.Visible = false;
            tlp.Controls.Add(num, 1, row);
            chk.CheckedChanged += (_, __) => { lbl.Visible = chk.Checked; num.Visible = chk.Checked; };
        }

        private IEnumerable<Control> GetAllControls(Control p)
        {
            foreach (Control c in p.Controls)
            {
                yield return c;
                foreach (var ch in GetAllControls(c))
                    yield return ch;
            }
        }

        private int ExcelSpalteZuIndex(string s)
        {
            int sum = 0;
            foreach (char c in s.ToUpper())
                sum = sum * 26 + (c - 'A' + 1);
            return sum;
        }

        private Excel.Workbook GetActiveWorkbookOrWarn(string context)
        {
            var app = Globals.ThisAddIn.Application as Excel.Application;
            var wb = (app != null) ? app.ActiveWorkbook : null;
            if (wb == null)
            {
                MessageBox.Show(
                    "Es ist aktuell keine Arbeitsmappe geöffnet.\n" +
                    "Bitte öffne deine MDL-Arbeitsmappe und klicke danach auf „Fragen neu laden“.",
                    context, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return wb;
        }

        private Excel.Worksheet FindSheet(Excel.Workbook wb, string name)
        {
            if (wb == null) return null;
            foreach (Excel.Worksheet s in wb.Worksheets)
                if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                    return s;
            return null;
        }

        private static List<string> ExpandMastTokens(string input)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            var tokens = input.Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                              .Select(t => t.Trim());

            var rx = new Regex(@"^(?<pre>\D*)(?<num>\d+)(?<suf>\D*)$");

            foreach (var tok in tokens)
            {
                var parts = tok.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 2)
                {
                    var m1 = rx.Match(parts[0]);
                    var m2 = rx.Match(parts[1]);

                    if (m1.Success && m2.Success)
                    {
                        string pre1 = m1.Groups["pre"].Value;
                        string suf1 = m1.Groups["suf"].Value;
                        string pre2 = m2.Groups["pre"].Value.Length > 0 ? m2.Groups["pre"].Value : pre1;
                        string suf2 = m2.Groups["suf"].Value.Length > 0 ? m2.Groups["suf"].Value : suf1;

                        if (pre1 != pre2 || suf1 != suf2)
                        {
                            result.Add(tok);
                            continue;
                        }

                        int start, end;
                        if (!int.TryParse(m1.Groups["num"].Value, out start) ||
                            !int.TryParse(m2.Groups["num"].Value, out end))
                        {
                            result.Add(tok);
                            continue;
                        }

                        if (end < start) { int tmp = start; start = end; end = tmp; }
                        int width = Math.Max(m1.Groups["num"].Value.Length, m2.Groups["num"].Value.Length);

                        for (int i = start; i <= end; i++)
                            result.Add($"{pre1}{i.ToString("D" + width)}{suf1}");

                        continue;
                    }
                }

                result.Add(tok);
            }

            return result.Distinct().ToList();
        }
    }
}