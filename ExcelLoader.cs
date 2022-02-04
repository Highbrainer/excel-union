using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace ExcelUnion
{
    class ExcelLoader
    {
        private string fileA;
        private int totalLines;

        public ExcelLoader(string a)
        {
            this.fileA = a;
        }

        private bool existsIn(List<object> row, List<List<object>> contentA)
        {
            foreach (List<object> candidate in contentA)
            {
                bool different = false;
                for (int i = 0; i < row.Count; ++i)
                {
                    if (row[i] == null)
                    {
                        if (candidate.Count < i)
                        {
                            continue;
                        }
                        if (candidate[i] != null)
                        {
                            different = true;
                            break;
                        }
                    }
                    else if (candidate.Count <= i)
                    {
                        throw new Exception("Nombre de colonnes différents ! (" + row.Count + " vs " + candidate.Count + ")");
                    }
                    else
                    if (!(row[i].Equals(candidate[i])))
                    {
                        different = true;
                        break;
                    }
                }
                if (!different)
                {
                    return true;
                }
            }
            return false;
        }

        public Content Content(string sheetname, int keyColumn, List<Column> chosenColumns, ProgressBar progress)
        {
            Content ret = new Content();
            ret.Sheet = sheetname;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            //int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(sheetname))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetname + "\" dans " + fileA);
                return ret;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            //progress.Maximum += rw;

            //lines
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                List<object> row = rCnt == 1 ? ret.Titles : new List<object>();
                string key = null;
                dynamic val = "!!!ERREUR!!!";
                try
                {
                    val = (range.Cells[rCnt, keyColumn + 1] as Excel.Range).Value2;
                    key = val/*.ToString()*/;
                }catch(Exception ex)
                {
                    MessageBox.Show("Dans l'onglet \"" + sheetname + "\" du fichier\n" + fileA + "\nle champ clef (colonne " + (keyColumn + 1) + ") de la ligne " + rCnt + " ne semble pas au bon format : '" + val + "'");
                }
                //TEMP
                //string c = (range.Cells[rCnt, 3] as Excel.Range).Value2;
                if (string.IsNullOrWhiteSpace(key))
                {
                    Console.WriteLine("Skipping line " + rCnt + " : " + key);
                    continue;
                }
                // FIN TEMP
                try
                {
                    if (rCnt > 1) ret.Lines.Add(key, row);
                    foreach (int cCnt in chosenColumns.Select(c => c.index + 1).ToList())
                    {
                        row.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                }
                catch (ArgumentException e)
                {
                    Console.WriteLine("Found doublon for " + key + " in " + fileA);
                }
                progress.Value += 1;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return ret;
        }

        internal static void GenerateUnion(string fileOut, Content contents1, Content contents2, ProgressBar progressBarA)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;


            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = xlWorkBook.Sheets.Add();

            int row = 0;
            {
                ++row;
                int col = 0;
                foreach (object o in contents1.Titles)
                {
                    xlWorkSheet.Cells[row, ++col] = o;
                }
                foreach (object o in contents2.Titles)
                {
                    xlWorkSheet.Cells[row, ++col] = o;
                }
                progressBarA.Value += 1;
            }

            foreach (string key in contents1.Lines.Keys)
            {
                ++row;
                int col = 0;
                foreach (object o in contents1.Lines[key])
                {
                    xlWorkSheet.Cells[row, ++col] = o;
                }
                if (contents2.Lines.ContainsKey(key))
                {
                    foreach (object o in contents2.Lines[key])
                    {
                        xlWorkSheet.Cells[row, ++col] = o;
                    }
                    progressBarA.Maximum -= 1;
                    contents2.Lines.Remove(key);
                }
                progressBarA.Value += 1;
            }

            int nbCols1 = 0;
            if (contents1.Lines.Values.Count > 0)
            {
                nbCols1 = contents1.Lines.Values.First().Count;
            } else
            {
                MessageBox.Show("L'onglet semble vide ! " + contents1.Sheet, "Erreur !", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                return;
            }

            foreach (string key in contents2.Lines.Keys)
            {
                ++row;
                int col = nbCols1;
                foreach (object o in contents2.Lines[key])
                {
                    xlWorkSheet.Cells[row, ++col] = o;
                }
                progressBarA.Value += 1;
            }

            xlWorkBook.SaveCopyAs(fileOut);

            xlWorkBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public int Lines(string sheetName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    if (sheet.Name.Equals(sheetName))
                    {
                        xlWorkSheet = sheet;
                    }
                }

                if (null == xlWorkSheet)
                {
                    MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + fileA);

                    return -1;
                }
                range = xlWorkSheet.UsedRange;
                return range.Rows.Count;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public List<Column> Columns(string sheetName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            List<Column> ret = new List<Column>();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    if (sheet.Name.Equals(sheetName))
                    {
                        xlWorkSheet = sheet;
                    }
                }

                if (null == xlWorkSheet)
                {
                    MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + fileA);
                    return ret;
                }

                range = xlWorkSheet.UsedRange;
                int cols = range.Columns.Count;
                for (int i = 0; i < cols; ++i)
                {
                    var title = (range.Cells[1, i + 1] as Excel.Range).Value2;
                    ret.Add(new Column(i, title));
                }
                return ret;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                if (null != xlWorkSheet) Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public List<string> Sheets()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            List<string> ret = new List<string>();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    ret.Add(sheet.Name);
                }
                return ret;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
