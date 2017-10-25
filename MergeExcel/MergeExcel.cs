using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    public class MergeExcel
    {
        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook bookDest = null;
        Microsoft.Office.Interop.Excel.Worksheet sheetDest = null;
        Microsoft.Office.Interop.Excel.Workbook bookSource = null;
        Microsoft.Office.Interop.Excel.Worksheet sheetSource = null;
        string[] _sourceFiles = null;
        string _destFile = string.Empty;
        string _columnEnd = string.Empty;
        int _headerRowCount = 0;
        int _currentRowCount = 0;
        public MergeExcel(string[] sourceFiles, string destFile, string columnEnd, int headerRowCount)
        {
            bookDest = (Microsoft.Office.Interop.Excel.Workbook)app.Workbooks.Add(Missing.Value);
            sheetDest = bookDest.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Microsoft.Office.Interop.Excel.Worksheet;
            sheetDest.Name = "Data";
            _sourceFiles = sourceFiles;
            _destFile = destFile;
            _columnEnd = columnEnd;
            _headerRowCount = headerRowCount;
        }
        void OpenBook(string fileName)
        {
            bookSource = app.Workbooks._Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            sheetSource = bookSource.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
        }
        void CloseBook()
        {
            bookSource.Close(false, Missing.Value, Missing.Value);
        }

        void CopyHeader()
        {
            Microsoft.Office.Interop.Excel.Range range = sheetSource.get_Range("A1", _columnEnd + _headerRowCount.ToString());
            range.Copy(sheetDest.get_Range("A1", Missing.Value));
            _currentRowCount += _headerRowCount;
        }
        void CopyData()
        {
            int sheetRowCount = sheetSource.UsedRange.Rows.Count;
            Microsoft.Office.Interop.Excel.Range range = sheetSource.get_Range(string.Format("A{0}", _headerRowCount), _columnEnd + sheetRowCount.ToString());
            range.Copy(sheetDest.get_Range(string.Format("A{0}", _currentRowCount), Missing.Value));
            _currentRowCount += range.Rows.Count;
        }
        void Save()
        {
            bookDest.Saved = true;
            bookDest.SaveCopyAs(_destFile);
        }
        void Quit()
        {
            app.Quit();
        }
        void DoMerge()
        {
            bool b = false;
            foreach (string strFile in _sourceFiles)
            {
                OpenBook(strFile);
                if (b == false)
                {
                    CopyHeader();
                    b = true;
                }
                CopyData();
                CloseBook();
            }
            Save();
            Quit();
        }
        public static System.Data.DataTable ViewData(String pathFile)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Static File From Base Path...........
            //Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "TestExcel.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //Dynamic File Using Uploader...........
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(pathFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
            Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

            string strCellData = "";
            double douCellData;
            int rowCnt = 0;
            int colCnt = 0;

            System.Data.DataTable dt = new System.Data.DataTable();
            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
            {
                string strColumn = "";
                strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                dt.Columns.Add(strColumn, typeof(string));
            }

            for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
            {
                string strData = "";
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    try
                    {
                        strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        strData += strCellData + "|";
                    }
                    catch (Exception ex)
                    {
                        douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        strData += douCellData.ToString() + "|";
                    }
                }
                strData = strData.Remove(strData.Length - 1, 1);
                dt.Rows.Add(strData.Split('|'));
            }

            excelBook.Close(true, null, null);
            excelApp.Quit();
            return dt;
        }
        public static void openExcel(String path)
        {
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);
            excel.Visible = true;

        }
        public static void DoMerge(string[] sourceFiles, string destFile, string columnEnd, int headerRowCount)
        {
            new MergeExcel(sourceFiles, destFile, columnEnd, headerRowCount).DoMerge();
        }
    }
}
