using System;
using System.Diagnostics;
using System.Globalization;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using WinForms = System.Windows.Forms;
using Microsoft.Win32;

namespace MDL_TEST_MA
{
    public partial class ThisAddIn
    {
        private const string RegKeyPath = @"Software\MDL_TEST_MA";
        private const string RegValueName = "WorkbookPath";

        private const string SheetMDL = "Master Document List";
        private const string SheetReporting = "Reporting";

        private const int FirstDataRow = 7; 
        private const int LastReportingCol = 37;

        private bool _eventsAttached = false;
        private EingabeForm _form;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            AttachAppEvents();

            try { EnsureWorkbookOpen(); } catch (Exception ex) { Debug.WriteLine("EnsureWorkbookOpen: " + ex); }

            try
            {
                if (_form == null || _form.IsDisposed)
                {
                    _form = new EingabeForm();
                    _form.Show();
                }
            }
            catch (Exception ex) { Debug.WriteLine("EingabeForm.Show: " + ex); }

            try { BuildReporting(); } catch (Exception ex) { Debug.WriteLine("BuildReporting at Startup: " + ex); }
            try { TryActivateMDL(); } catch (Exception ex) { Debug.WriteLine("TryActivateMDL at Startup: " + ex); }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            DetachAppEvents();
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        private void AttachAppEvents()
        {
            if (_eventsAttached) return;
            this.Application.SheetChange += Application_SheetChange;
            this.Application.WorkbookOpen += Application_WorkbookOpen;
            _eventsAttached = true;
        }

        private void DetachAppEvents()
        {
            if (!_eventsAttached) return;
            this.Application.SheetChange -= Application_SheetChange;
            this.Application.WorkbookOpen -= Application_WorkbookOpen;
            _eventsAttached = false;
        }

        private void EnsureWorkbookOpen()
        {
            var app = this.Application as Excel.Application;
            if (app == null) return;

            if (app.Workbooks.Count > 0) return;

            string saved = LoadWorkbookPathFromRegistry();
            if (!string.IsNullOrEmpty(saved) && File.Exists(saved))
            {
                try { app.Workbooks.Open(saved); return; }
                catch (Exception ex) { Debug.WriteLine("Open saved workbook failed: " + ex); }
            }

            using (var ofd = new WinForms.OpenFileDialog
            {
                Title = "MDL-Arbeitsmappe auswählen",
                Filter = "Excel-Dateien|*.xlsx;*.xlsm;*.xls|Alle Dateien|*.*"
            })
            {
                if (ofd.ShowDialog() == WinForms.DialogResult.OK)
                {
                    try
                    {
                        SaveWorkbookPathToRegistry(ofd.FileName);
                        app.Workbooks.Open(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        WinForms.MessageBox.Show(
                            "Die Arbeitsmappe konnte nicht geöffnet werden:\n" + ex.Message,
                            "Öffnen fehlgeschlagen",
                            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    }
                }
                else
                {
                    WinForms.MessageBox.Show(
                        "Es wurde keine Arbeitsmappe ausgewählt. Du kannst sie später über Excel öffnen.\n" +
                        "In der Eingabemaske klickst du dann auf „Fragen neu laden“.",
                        "Hinweis", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
            }
        }

        private string LoadWorkbookPathFromRegistry()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath, false))
                    return key?.GetValue(RegValueName) as string;
            }
            catch (Exception ex) { Debug.WriteLine("LoadWorkbookPathFromRegistry: " + ex); return null; }
        }

        private void SaveWorkbookPathToRegistry(string path)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegKeyPath))
                    key?.SetValue(RegValueName, path);
            }
            catch (Exception ex) { Debug.WriteLine("SaveWorkbookPathToRegistry: " + ex); }
        }

        public void ChangeWorkbookPathInteractive()
        {
            using (var ofd = new WinForms.OpenFileDialog
            {
                Title = "Neue MDL-Arbeitsmappe auswählen",
                Filter = "Excel-Dateien|*.xlsx;*.xlsm;*.xls|Alle Dateien|*.*"
            })
            {
                if (ofd.ShowDialog() == WinForms.DialogResult.OK)
                {
                    SaveWorkbookPathToRegistry(ofd.FileName);
                    WinForms.MessageBox.Show("Neuer Pfad gespeichert:\n" + ofd.FileName, "OK");
                }
            }
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            try { TryActivateMDL(Wb); } catch (Exception ex) { Debug.WriteLine("WorkbookOpen/TryActivateMDL: " + ex); }
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            try
            {
                var sheet = Sh as Excel.Worksheet;
                if (sheet == null) return;
                if (!string.Equals(sheet.Name, SheetMDL, StringComparison.OrdinalIgnoreCase)) return;

                foreach (Excel.Range cell in Target.Cells)
                {
                    int col = cell.Column;
                    int row = cell.Row;

                    if (col == 4 || col == 5)
                    {
                        var expCell = (Excel.Range)sheet.Cells[row, 4];
                        var actCell = (Excel.Range)sheet.Cells[row, 5];

                        int expectedWeek = 0;
                        DateTime actualDate = default(DateTime);
                        bool okWeek = TryGetIntFromValue2(expCell.Value2, out expectedWeek);
                        bool okDate = TryGetDateFromValue2(actCell.Value2, out actualDate);

                        if (okWeek && okDate)
                        {
                            int actualWeek = GetIsoWeek(actualDate);
                            int oleColor = expectedWeek == actualWeek
                                ? ColorTranslator.ToOle(Color.LightGreen)
                                : ColorTranslator.ToOle(Color.Orange);
                            expCell.Interior.Color = oleColor;
                            actCell.Interior.Color = oleColor;
                        }

                        ReleaseCom(expCell);
                        ReleaseCom(actCell);
                    }

                    if (col >= 33 && col <= 37)
                    {
                        try { ColorizeRowStatus(sheet, row); } catch (Exception ex) { Debug.WriteLine("ColorizeRowStatus: " + ex); }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine("Application_SheetChange: " + ex); }
        }

        public void BuildReporting()
        {
            var app = this.Application as Excel.Application;
            if (app == null) return;

            Excel.Worksheet wsMDL = FindSheet(app, SheetMDL);
            if (wsMDL == null) return;

            Excel.Worksheet wsRep = FindSheet(app, SheetReporting);
            if (wsRep == null)
            {
                wsRep = (Excel.Worksheet)app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
                wsRep.Name = SheetReporting;
            }

            bool oldEvents = app.EnableEvents;
            bool oldScreen = app.ScreenUpdating;
            var oldCalc = app.Calculation;

            Excel.Range used = null;
            Excel.ChartObjects charts = null;

            try
            {
                app.EnableEvents = false;
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                try { ApplyStatusColoringToAllRows(wsMDL); } catch (Exception ex) { Debug.WriteLine("ApplyStatusColoring: " + ex); }

                used = wsMDL.UsedRange;
                int lastRow = used.Row + used.Rows.Count - 1;
                if (lastRow < FirstDataRow) lastRow = FirstDataRow - 1;

                wsRep.Cells.Clear();

                try
                {
                    charts = (Excel.ChartObjects)wsRep.ChartObjects(Type.Missing);
                    foreach (Excel.ChartObject co in charts) { co.Delete(); ReleaseCom(co); }
                }
                catch (Exception ex) { Debug.WriteLine("Delete charts: " + ex); }

                if (lastRow >= FirstDataRow)
                {
                    Excel.Range rng = wsMDL.Range[wsMDL.Cells[FirstDataRow, 1], wsMDL.Cells[lastRow, LastReportingCol]];
                    object valObj = rng.Value2;

                    int totalFilledRows = 0;
                    int uploaded = 0;

                    int cntFormalRejected = 0;
                    int cntFachlichRejected = 0;
                    int cntFreigegeben = 0;
                    int cntFormalChecked = 0;
                    int cntFachlichChecked = 0;
                    int cntSonstige = 0;

                    if (valObj is object[,] vals)
                    {
                        int rCount = vals.GetLength(0);
                        int cCount = vals.GetLength(1);

                        for (int r = 1; r <= rCount; r++)
                        {
                            bool anyFilled = false;
                            for (int c = 1; c <= Math.Min(32, cCount); c++)
                            {
                                if (vals[r, c] != null && vals[r, c].ToString().Trim().Length > 0)
                                { anyFilled = true; break; }
                            }
                            if (anyFilled) totalFilledRows++;

                            bool isUploaded = (5 <= cCount) && vals[r, 5] != null && vals[r, 5].ToString().Trim().Length > 0;
                            if (!isUploaded) continue;
                            uploaded++;

                            string ag = GetCellStr(vals, r, 33);
                            string ah = GetCellStr(vals, r, 34);
                            string ai = GetCellStr(vals, r, 35);
                            string aj = GetCellStr(vals, r, 36);
                            string ak = GetCellStr(vals, r, 37);

                            bool AGok = HasCheck(ag);
                            bool AHok = HasCheck(ah);
                            bool AIx = HasCross(ai);
                            bool AJx = HasCross(aj);
                            bool AKok = HasCheck(ak);

                            bool aiEmpty = IsEmpty(ai);
                            bool ajEmpty = IsEmpty(aj);
                            bool akEmpty = IsEmpty(ak);
                            bool agEmpty = IsEmpty(ag);
                            bool ahEmpty = IsEmpty(ah);

                            if (AIx) { cntFormalRejected++; continue; }
                            if (AJx) { cntFachlichRejected++; continue; }
                            if (AKok || (AGok && AHok && aiEmpty && ajEmpty && akEmpty))
                            { cntFreigegeben++; continue; }
                            if (AGok && ahEmpty && aiEmpty && ajEmpty && akEmpty)
                            { cntFormalChecked++; continue; }
                            if (AHok && agEmpty && aiEmpty && ajEmpty && akEmpty)
                            { cntFachlichChecked++; continue; }

                            cntSonstige++;
                        }
                    }

                    int outstanding = Math.Max(0, totalFilledRows - uploaded);

                    wsRep.Cells[1, 1].Value2 = "Reporting";
                    wsRep.Cells[2, 1].Value2 = "Dokumentenanzahl insgesamt:";
                    wsRep.Cells[2, 2].Value2 = totalFilledRows;

                    wsRep.Columns["A:B"].ColumnWidth = 35;
                    wsRep.Columns["D:E"].ColumnWidth = 40;

                    wsRep.Cells[4, 1].Value2 = "Kreisdiagramm 1: Gesamtübersicht";
                    wsRep.Cells[5, 1].Value2 = "Kategorie";
                    wsRep.Cells[5, 2].Value2 = "Wert";
                    wsRep.Cells[6, 1].Value2 = "hochgeladene Dokumente";
                    wsRep.Cells[6, 2].Value2 = uploaded;
                    wsRep.Cells[7, 1].Value2 = "ausstehende Dokumente";
                    wsRep.Cells[7, 2].Value2 = outstanding;

                    Excel.Range data1 = wsRep.Range[wsRep.Cells[6, 1], wsRep.Cells[7, 2]];

                    charts = (Excel.ChartObjects)wsRep.ChartObjects(Type.Missing);
                    Excel.Range anchor1 = wsRep.Range["A9"];
                    Excel.ChartObject ch1 = charts.Add(anchor1.Left, anchor1.Top, 420, 300);
                    ch1.Name = "MDL_Pie_Overview";
                    Excel.Chart chart1 = ch1.Chart;
                    chart1.ChartType = Excel.XlChartType.xlPie;
                    chart1.SetSourceData(data1);
                    chart1.HasTitle = true;
                    chart1.ChartTitle.Text = "Dokumentenübersicht (gesamt: " + totalFilledRows + ")";
                    chart1.HasLegend = true;

                    Excel.Series s1 = (Excel.Series)chart1.SeriesCollection(1);
                    s1.HasDataLabels = true;
                    var dl1 = s1.DataLabels();
                    dl1.ShowCategoryName = true;
                    dl1.ShowValue = true;
                    dl1.ShowPercentage = true;
                    dl1.Separator = "\n";

                    wsRep.Cells[4, 4].Value2 = "Kreisdiagramm 2: Status hochgeladener Dokumente";
                    wsRep.Cells[5, 4].Value2 = "Kategorie";
                    wsRep.Cells[5, 5].Value2 = "Wert";

                    if (uploaded > 0)
                    {
                        int row = 6;

                        wsRep.Cells[row, 4].Value2 = "formal zurückgewiesen";
                        wsRep.Cells[row, 5].Value2 = cntFormalRejected; row++;

                        wsRep.Cells[row, 4].Value2 = "fachlich zurückgewiesen";
                        wsRep.Cells[row, 5].Value2 = cntFachlichRejected; row++;

                        wsRep.Cells[row, 4].Value2 = "freigegeben";
                        wsRep.Cells[row, 5].Value2 = cntFreigegeben; row++;

                        wsRep.Cells[row, 4].Value2 = "formal geprüft";
                        wsRep.Cells[row, 5].Value2 = cntFormalChecked; row++;

                        wsRep.Cells[row, 4].Value2 = "fachlich geprüft";
                        wsRep.Cells[row, 5].Value2 = cntFachlichChecked; row++;

                        if (cntSonstige > 0)
                        {
                            wsRep.Cells[row, 4].Value2 = "sonstige";
                            wsRep.Cells[row, 5].Value2 = cntSonstige; row++;
                        }

                        Excel.Range data2 = wsRep.Range[wsRep.Cells[6, 4], wsRep.Cells[row - 1, 5]];

                        Excel.Range anchor2 = wsRep.Range["H9"];
                        Excel.ChartObject ch2 = charts.Add(anchor2.Left, anchor2.Top, 620, 360);
                        ch2.Name = "MDL_Pie_Uploaded";
                        Excel.Chart chart2 = ch2.Chart;
                        chart2.ChartType = Excel.XlChartType.xlPie;
                        chart2.SetSourceData(data2);
                        chart2.HasTitle = true;
                        chart2.ChartTitle.Text = "Status hochgeladener Dokumente (n=" + uploaded + ")";
                        chart2.HasLegend = true;

                        Excel.Series s2 = (Excel.Series)chart2.SeriesCollection(1);
                        s2.HasDataLabels = true;

                        var dl2 = s2.DataLabels();
                        dl2.ShowCategoryName = false;
                        dl2.ShowValue = false;
                        dl2.ShowPercentage = true;
                        dl2.NumberFormat = "0%";
                        dl2.Separator = "\n";
                        dl2.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd;
                        s2.HasLeaderLines = true;
                        try { dl2.Font.Size = 9; } catch { }

                        try
                        {
                            var grp2 = (Excel.ChartGroup)chart2.ChartGroups(1);
                            grp2.FirstSliceAngle = 15;
                            ReleaseCom(grp2);
                        }
                        catch { }

                        ReleaseCom(dl2);
                        ReleaseCom(s2);
                        ReleaseCom(chart2);
                        ReleaseCom(ch2);
                        ReleaseCom(data2);
                        ReleaseCom(anchor2);
                    }
                    else
                    {
                        wsRep.Cells[6, 4].Value2 = "Keine hochgeladenen Dokumente vorhanden.";
                    }

                    ReleaseCom(dl1);
                    ReleaseCom(s1);
                    ReleaseCom(chart1);
                    ReleaseCom(ch1);
                    ReleaseCom(data1);
                    ReleaseCom(anchor1);
                    ReleaseCom(rng);
                }
                else
                {
                    wsRep.Cells[1, 1].Value2 = "Keine Daten in 'Master Document List' (ab Zeile 7).";
                }
            }
            finally
            {
                app.Calculation = oldCalc;
                app.ScreenUpdating = oldScreen;
                app.EnableEvents = oldEvents;

                ReleaseCom(charts);
                ReleaseCom(used);
            }
        }

        private void TryActivateMDL(Excel.Workbook wb = null)
        {
            var app = this.Application as Excel.Application;
            if (wb == null && app != null) wb = app.ActiveWorkbook;
            if (wb == null) return;

            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                if (string.Equals(s.Name, SheetMDL, StringComparison.OrdinalIgnoreCase))
                {
                    var ws = s as Excel._Worksheet;
                    ws?.Activate();
                    break;
                }
            }
        }

        private Excel.Worksheet FindSheet(Excel.Application app, string name)
        {
            try
            {
                foreach (Excel.Worksheet s in app.Worksheets)
                    if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                        return s;
            }
            catch (Exception ex) { Debug.WriteLine("FindSheet: " + ex); }
            return null;
        }

        private static bool IsEmpty(string s) => string.IsNullOrWhiteSpace(s);
        private static bool HasCheck(string s) => !string.IsNullOrEmpty(s) && (s.Contains("✔") || s.Contains("✓"));
        private static bool HasCross(string s) => !string.IsNullOrEmpty(s) &&
                                                 (s.Contains("❌") || s.Contains("✖") || s.Equals("x", StringComparison.OrdinalIgnoreCase));

        private static string GetCellStr(object[,] vals, int r, int c)
        {
            if (vals == null) return "";
            object o = vals[r, c];
            return o == null ? "" : o.ToString();
        }

        private static bool TryGetIntFromValue2(object value2, out int result)
        {
            if (value2 == null) { result = 0; return false; }
            if (value2 is double d) { result = (int)Math.Round(d); return true; }
            if (value2 is int i) { result = i; return true; }
            var s = Convert.ToString(value2);
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out result)
                || int.TryParse(s, out result);
        }

        private static bool TryGetDateFromValue2(object value2, out DateTime dt)
        {
            if (value2 == null) { dt = default(DateTime); return false; }
            if (value2 is double d) { dt = DateTime.FromOADate(d); return true; }
            if (value2 is DateTime dd) { dt = dd; return true; }
            var s = Convert.ToString(value2);
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return true;
            if (DateTime.TryParse(s, out dt)) return true;
            return false;
        }

        private static int GetIsoWeek(DateTime date)
        {
            var cal = CultureInfo.InvariantCulture.Calendar;
            return cal.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private static void ReleaseCom(object com)
        {
            if (com == null) return;
            try { Marshal.FinalReleaseComObject(com); } catch { }
        }

        private void ColorizeRowStatus(Excel.Worksheet sheet, int row)
        {
            if (row < FirstDataRow) return;

            Excel.Range rng = null;
            try
            {
                rng = sheet.Range[sheet.Cells[row, 33], sheet.Cells[row, 37]];
                string ai = sheet.Cells[row, 35].Value2?.ToString() ?? "";
                string aj = sheet.Cells[row, 36].Value2?.ToString() ?? "";
                string ak = sheet.Cells[row, 37].Value2?.ToString() ?? "";

                bool reject = HasCross(ai) || HasCross(aj);
                bool released = HasCheck(ak);

                if (reject)
                {
                    rng.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    rng.Interior.Color = ColorTranslator.ToOle(Color.Red);
                }
                else if (released)
                {
                    rng.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    rng.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                }
                else
                {
                    rng.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                }
            }
            finally
            {
                ReleaseCom(rng);
            }
        }

        private void ApplyStatusColoringToAllRows(Excel.Worksheet wsMDL)
        {
            Excel.Range used = null;
            try
            {
                used = wsMDL.UsedRange;
                int lastRow = used.Row + used.Rows.Count - 1;
                if (lastRow < FirstDataRow) return;

                for (int r = FirstDataRow; r <= lastRow; r++)
                {
                    try { ColorizeRowStatus(wsMDL, r); } catch (Exception ex) { Debug.WriteLine("ColorizeRowStatus(all): " + ex); }
                }
            }
            finally
            {
                ReleaseCom(used);
            }
        }
    }
}