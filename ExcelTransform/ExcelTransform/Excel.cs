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
        string path = "";
        Workbook WB;
        Worksheet WS;
        string[,] content;
        public int rows;
        public int cols;
        public int specfialCol;

        public Excel(string path, int sheet)
        {
            this.path = path;
            WB = excel.Workbooks.Open(path);
            WS = WB.Worksheets[sheet];
            this.rows = this.WS.UsedRange.Rows.Count;
            this.cols = this.WS.UsedRange.Columns.Count;
            this.content = new string[this.WS.UsedRange.Rows.Count, this.WS.UsedRange.Columns.Count];
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
            string _expression = @"[\d]{1,4}([.,][\d]{0,10})?";
            string _expression_symbol = @"^[\u4e00-\u9fa5]{0,100}$";
           
            Regex _objNotNumberPattern = new Regex(_expression);
            Regex _objNotSymbolPattern = new Regex(_expression_symbol);
            Match _match , _match_2;

            //copy percentage out
            for (int i = 1; i <= this.rows; i++)
            {

                _match = _objNotNumberPattern.Match(Convert.ToString(WS.Cells[i, specfialCol].Value));
                int _temp_col = specfialCol + 1;
                if (_match.Success)
                {
                    string _string = _match.Value;
                    WriteToCell(i, _temp_col, _match.Value + " %");
                }
                else
                {
                    WriteToCell(i, _temp_col, "");
                }

            }

            //remove original cell
            for (int i = 1; i <= this.rows; i++)
            {

                _match_2 = _objNotNumberPattern.Match(Convert.ToString(WS.Cells[i, specfialCol].Value2));
                string[] pattern = new string[] { "：", ':'.ToString(), '*'.ToString(), '.'.ToString(), '"'.ToString() };
                char[] array = string.Join(string.Empty, pattern).ToCharArray();

                if (_match_2.Success)
                {
                    string _string = Convert.ToString(WS.Cells[i, specfialCol].Value2);
                    _string = _string.Replace("%", "").Replace(_match_2.Value, "").Trim(array);

                    WriteToCell(i, specfialCol, _string);
                }
                Court.PROCESSED_ROWS += 1;
                Court.LABEL.Text = "Done: " + Court.PROCESSED_ROWS;
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
    }
}
