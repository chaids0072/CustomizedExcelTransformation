using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ExcelTransform
{
    class Excel
    {
        _Application excel = new _Excel.Application();
        
        public Workbook WB { get; private set; }
        public Worksheet WS { get; private set; }
        public  string[,] content { get; private set; }
        public int rows { get; private set; }
        public int cols { get; private set; }
        public int specfialCol;
        public string path { get; private set; }
        public static int processedRows;

        public Excel(string path, int sheet)
        {
            this.path = path;
            this.WB = excel.Workbooks.Open(path);
            this.WS = WB.Worksheets[sheet];
            this.rows = this.WS.UsedRange.Rows.Count;
            this.cols = this.WS.UsedRange.Columns.Count;
            this.content = new string[this.WS.UsedRange.Rows.Count, this.WS.UsedRange.Columns.Count];
            Excel.processedRows = 0;
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (WS.Cells[i, j].Value2 != null)
            {
                return WS.Cells[i, j].Value2;
            }
            else
            {
                return "";
            }
        }

        public string[,] ReadAll()
        {
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    if (WS.Cells[i, j].Value2 != null)
                    {
                        int _tempX = i - 1;
                        int _tempY = j - 1;
                        content[_tempX, _tempY] = Convert.ToString(WS.Cells[i, j].Value2);
                    }
                }
            }
            return this.content;
        }

        public void pushCol(int col)
        {
            for (int i = WS.UsedRange.Columns.Count ; i > col; i--)
            {
                for (int v = WS.UsedRange.Row; v < this.rows + WS.UsedRange.Row; v++)
                {
                    WriteToCell(v,i+1, WS.Cells[v, i].Value2);
                }
            }
        }

        public void pushSingleCol(int row, int col)
        {
            for (int i = WS.UsedRange.Columns.Count; i > col; i--)
            {
                WriteToCell(row, i + 1, WS.Cells[row, i].Value2);
            }
        }

        public void WriteToCell(int i, int j, string s) {
            WS.Cells[i, j] = s;
        }

        public void UpdateMatrix()
        {
            this.rows = this.WS.UsedRange.Rows.Count;
            this.cols = this.WS.UsedRange.Columns.Count;
        }

        public void Save()
        {
            WB.Save();
        }

        public void SaveAs(string path)
        {
            WB.SaveAs(path);
        }

        public void Close()
        {
            WB.Close();
        }

        public void Clear()
        {
            processedRows = 0;
        }

        public void columnDelete(int column)
        {
            WS.Cells[1, column].EntireColumn.Delete(null);
        }
    }
}
