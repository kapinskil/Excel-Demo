using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Demo
{
    class Excel
    {
        string _path = "";
        _Application _excel = new _Excel.Application();
        Workbook _workbook;
        Worksheet _worksheet;

        public Excel()
        {

        }

        public Excel(string path, int sheet)
        {
            _path = path;
            _workbook = _excel.Workbooks.Open(_path);
            _worksheet = (Worksheet)_workbook.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            var result = default(string);
            var sno = (_Excel.Range)_worksheet.Cells[i, j];

            if(sno.Value2 != null)
            {
                result = sno.Value2;
            }

            return result;
        }

        public void WriteToCell(int i, int j, string value)
        {
            var sno = (_Excel.Range)_worksheet.Cells[i, j];
            sno.Value2 = value;
        }

        public void Save()
        {
            _workbook.Save();
        }

        public void SaveAs(string path)
        {
            _workbook.SaveAs2(path);
        }

        public void Close()
        {
            _workbook.Close();
        }

        public void CreateNewWorkbook()
        {
            _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _worksheet = _workbook.Worksheets[1];
        }
        public void CreateNewWorksheet()
        {
            var temptSheet = _workbook.Worksheets.Add(After: _worksheet);
        }

        public void SelectWorksheet(int sheet)
        {
            _worksheet = _workbook.Worksheets[1];
        }

        public void DeleteWorksheet(int sheet)
        {
            var temp = (Worksheet)_workbook.Worksheets[sheet];
            temp.Delete();
        }
    }
}
