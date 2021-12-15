using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

using System.Data.SQLite;


namespace ExcelUnion
{
    class ExcelComparatorSql
    {
        private string fileA;
        private string fileB;
        private string sheetName;

        public ExcelComparatorSql(string sheetName, string a, string b)
        {
            this.fileA = a;
            this.fileB = b;
            this.sheetName = sheetName;

        }

        private void LoadTable(string tablename, string filename, ProgressBar progress)
        {

            string cs = "Data Source=:memory:";
            string stm = "SELECT SQLITE_VERSION()";

            using var con = new SQLiteConnection(cs);
            con.Open();

            using var cmd = new SQLiteCommand(stm, con);
            string version = cmd.ExecuteScalar().ToString();

            Console.WriteLine($"SQLite version: {version}");


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            List<List<object>> content = new List<List<object>>();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(this.sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + filename);

                return;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            progress.Maximum += rw;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                List<object> row = new List<object>();
                content.Add(row);
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    row.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                }
                progress.Value += 1;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return;
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

        public void OnlyInA(ProgressBar progress)
        {
            OnlyIn(fileB, fileA, progress);
        }

        public void OnlyInB(ProgressBar progress)
        {
            OnlyIn(fileA, fileB, progress);
        }

        public void OnlyIn(string fileA, string fileB, ProgressBar progress)
        {
            List<List<object>> contentA = Content(fileA, progress);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            List<List<object>> content = new List<List<object>>();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileB, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(this.sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + fileB);
                return;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            progress.Maximum += rw;
            

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                List<object> row = new List<object>();
                content.Add(row);
                bool foundDifference = false;
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    var cell = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    row.Add(cell);
                }

                if (existsIn(row, contentA))
                {
                    // same line remove it !
                    for (cCnt = 1; cCnt <= cl; cCnt++)
                    {
                        (range.Cells[rCnt, cCnt] as Excel.Range).Value2 = "";
                        //range[rCnt, cCnt].Delete();
                    }
                }

                progress.Value += 1;
            }

            xlWorkBook.SaveCopyAs(Path.GetDirectoryName(fileB) + @"\Dans-" + Path.GetFileName(fileB) + "-mais-pas-dans-" + Path.GetFileName(fileA));

            xlWorkBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private List<List<object>> Content(string file, ProgressBar progress)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            List<List<object>> content = new List<List<object>>();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(this.sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + file);

                return content;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            progress.Maximum += rw;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                List<object> row = new List<object>();
                content.Add(row);
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    row.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                }
                progress.Value += 1;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return content;
        }
    }
}
