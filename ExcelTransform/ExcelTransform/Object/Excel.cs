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
        public int specfialCol { get; private set; }
        public string path { get; private set; }
        public static int processedRows { get; private set; }

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
            Console.WriteLine(rows + " + " + cols);
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
            int _new_col = col + 1;
            for (int i = this.cols + 1; i > _new_col; i--)
            {
                for (int j = 1; j <= this.rows; j++)
                {
                    int _temp_col = i - 1;
                    WS.Cells[j, i].Value2 = WS.Cells[j, _temp_col].Value2;
                }
            }
        }

        public void specialReplace()
        {
            string expression = @"[\d]{1,4}([.,][\d]{0,10})?";
            string expression_symbol = @"^[\u4e00-\u9fa5]{0,100}$";
           
            Regex objNotNumberPattern = new Regex(expression);
            Regex objNotSymbolPattern = new Regex(expression_symbol);
            Match match , match2;

            //copy percentage out
            for (int i = 1; i <= this.rows; i++)
            {

                match = objNotNumberPattern.Match(Convert.ToString(WS.Cells[i, specfialCol].Value));
                int tempCol = specfialCol + 1;
                if (match.Success)
                {
                    string _string = match.Value;
                    WriteToCell(i, tempCol, match.Value + " %");
                }
                else
                {
                    WriteToCell(i, tempCol, "");
                }

            }

            //remove original cell
            for (int i = 1; i <= this.rows; i++)
            {

                match2 = objNotNumberPattern.Match(Convert.ToString(WS.Cells[i, specfialCol].Value2));
                string[] pattern = new string[] { "：", ':'.ToString(), '*'.ToString(), '.'.ToString(), '"'.ToString() };
                char[] array = string.Join(string.Empty, pattern).ToCharArray();

                if (match2.Success)
                {
                    string tempString = Convert.ToString(WS.Cells[i, specfialCol].Value2);
                    tempString = tempString.Replace("%", "").Replace(match2.Value, "").Trim(array);

                    WriteToCell(i, specfialCol, tempString);
                }
                Excel.processedRows += 1;
                UpdateFunction.UpdateProcessedRow();
            }
        }

        private string KeepChinese(string str)
        {
            string chineseString = "";

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] >= 0x4E00 && str[i] <= 0x9FA5)
                {
                    chineseString += str[i];
                }
            }

            return chineseString;
        }

        public void WriteToCell(int i, int j, string s) {
            WS.Cells[i, j].Value2 = s;
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
            Excel.processedRows = 0;
        }
    }
}
