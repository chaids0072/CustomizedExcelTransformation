using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTransform
{
    class UpdateExcel
    {
        public static void UpdateProcessedRow()
        {
            Court.LABEL.Text = "Done: " + Excel.processedRows;
        }

        public static void InitializeComponent()
        {
            UpdateProcessedRow();
        }

        public static int UpdateSpecialColumn()
        {
            int colToReturn = 0;
            for (int row = Court.excel.WS.UsedRange.Row; row < Court.excel.rows + Court.excel.WS.UsedRange.Row; row++)
            {
                for (int col = Court.excel.WS.UsedRange.Column; col < Court.excel.cols + Court.excel.WS.UsedRange.Column; col++)
                {
                    string[] lines = Convert.ToString(Court.excel.WS.Cells[row, col].Value).Split(
                        new[] { "\r\n", "\r", "\n" },
                        StringSplitOptions.None
                    );

                    if (lines.Length > 2)
                    {
                        colToReturn = col;
                        return colToReturn;
                    }
                }
            }

            return colToReturn;
        }

        public static int UpdateSpecialColumnByRow(int row)
        {
            int colToReturn = 0;
            for (int col = Court.excel.WS.UsedRange.Column; col < Court.excel.cols + Court.excel.WS.UsedRange.Column; col++)
            {
                if (Court.excel.WS.Cells[row, col].Value == null) {
                    return colToReturn;
                }

                string[] lines = Convert.ToString(Court.excel.WS.Cells[row, col].Value).Split(
                    new[] { "\r\n", "\r", "\n" },
                    StringSplitOptions.None
                );

                if (lines.Length > 2)
                {
                    colToReturn = col;
                    return colToReturn;
                }
            }

            return colToReturn;
        }

        public static void UpdateSpecialSeparate()
        {
            ArrayList myLineList = new ArrayList();
            ArrayList myNumberList = new ArrayList();
            int i = Court.excel.WS.UsedRange.Row;
            int lineCounter = Court.excel.WS.UsedRange.Row;
            while(i < Court.excel.rows + Court.excel.WS.UsedRange.Row)
            {
                myNumberList.Add(lineCounter);
                //getting all the names in this cell
                string[] lines = Convert.ToString(Court.excel.WS.Cells[i, Court.excel.specfialCol].Value).Split(
                    new[] { "\r\n", "\r", "\n" },
                    StringSplitOptions.None
                );
                lineCounter = lineCounter + lines.Length;
                myLineList.Add(lines);
                i++;
            }

            for (int x = 0; x < myLineList.Count; x++)
            {
                string[] tempString = (string[]) myLineList[x];
                int tempLine = (int)myNumberList[x];

                for (int y = 0; y < tempString.Length; y++)
                {
                   Court.excel.WriteToCell((tempLine + y), (Court.excel.specfialCol + 1), tempString[y]);
                }
            }

            //Copy row number
            int lastCopyRow = Court.excel.rows + Court.excel.WS.UsedRange.Row - 1;
            for (int iRow = myNumberList.Count - 1; iRow >= 0; iRow--)
            {
                for (int v = Court.excel.WS.UsedRange.Column; v <= Court.excel.cols + Court.excel.WS.UsedRange.Column; v++)
                {
                    if (v != Court.excel.specfialCol && v != Court.excel.specfialCol + 1)
                    {
                        Court.excel.WriteToCell((int)myNumberList[iRow], v, Convert.ToString(Court.excel.WS.Cells[lastCopyRow, v].Value));

                        //Clear copied rows
                        if ((int)myNumberList[iRow] != lastCopyRow)
                        {
                            Court.excel.WS.Cells[lastCopyRow, v].Clear();
                        }
                    } 
                }
                lastCopyRow--;
            }

            //Merge row now
            for (int m = myNumberList.Count - 1; m >= 0; m--)
            {
                for (int v = Court.excel.WS.UsedRange.Column; v <= Court.excel.cols + Court.excel.WS.UsedRange.Column; v++)
                {
                    if (v != Court.excel.specfialCol && v != Court.excel.specfialCol + 1)
                    {
                        Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], v], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, v]].Merge();
                        Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], v], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, v]].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], v], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, v]].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], v], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, v]].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], v], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, v]].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }
                }

                //Court.excel.WS.Range[Court.excel.WS.Cells[(int)myNumberList[m], col], Court.excel.WS.Cells[(int)myNumberList[m] + ((string[])myLineList[m]).Length - 1, col]].Merge();
                Excel.processedRows += 1;
                UpdateProcessedRow();
            }

            //delete special row
            Court.excel.columnDelete(Court.excel.specfialCol);
            Court.excel.WS.Columns.AutoFit();
            Court.excel.WS.Rows.AutoFit();
        }

        public static void UpdateSpecialSeparateRandom()
        {
            int i = Court.excel.WS.UsedRange.Row;
            while (i < Court.excel.rows + Court.excel.WS.UsedRange.Row)
            {
                int specialCol = UpdateSpecialColumnByRow(i);
                if (specialCol != 0)
                {
                    //add new rows below the special row
                    string[] lines = Convert.ToString(Court.excel.WS.Cells[i, specialCol].Value).Split(
                    new[] { "\r\n", "\r", "\n" },
                    StringSplitOptions.None
                    );

                    for (int j = 1; j < lines.Length; j++)
                    {
                        Court.excel.WS.Rows[i + 1].Insert();
                    }

                    for (int j = 1; j < lines.Length; j++)
                    {
                        Court.excel.WriteToCell(i + j, specialCol, lines[j]);
                    }

                    Court.excel.WriteToCell(i, specialCol, lines[0]);

                    //Merge Rows
                    for (int v = Court.excel.WS.UsedRange.Column; v < Court.excel.cols + Court.excel.WS.UsedRange.Column; v++)
                    {
                        if (v != specialCol)
                        {
                            Court.excel.WS.Range[Court.excel.WS.Cells[i, v], Court.excel.WS.Cells[i + lines.Length - 1, v]].Merge();
                            Court.excel.WS.Range[Court.excel.WS.Cells[i, v], Court.excel.WS.Cells[i + lines.Length - 1, v]].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            Court.excel.WS.Range[Court.excel.WS.Cells[i, v], Court.excel.WS.Cells[i + lines.Length - 1, v]].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        }
                        else
                        {
                            Court.excel.WS.Range[Court.excel.WS.Cells[i, v], Court.excel.WS.Cells[i + lines.Length - 1, v]].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            Court.excel.WS.Range[Court.excel.WS.Cells[i, v], Court.excel.WS.Cells[i + lines.Length - 1, v]].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        }
                    }

                    i += lines.Length;
                }
                else
                {
                    i += 1;
                }
                Court.excel.UpdateMatrix();
                Excel.processedRows += 1;
                UpdateProcessedRow();
            }

            Court.excel.WS.Columns.AutoFit();
            Court.excel.WS.Rows.AutoFit();
        }
        public static void UpdateSpecialReplace()
        {
            string expression = Court.currentExtraction;
            string expression_symbol = Court.currentExtractionSymbol;

            Regex objNotNumberPattern = new Regex(expression, RegexOptions.RightToLeft);
            Regex objNotSymbolPattern = new Regex(expression_symbol);
            Match match, match2;

            //copy percentage out
            for (int i = Court.excel.WS.UsedRange.Row; i < Court.excel.rows + Court.excel.WS.UsedRange.Row; i++)
            {
                match = objNotNumberPattern.Match(Convert.ToString(Court.excel.WS.Cells[i, Court.excel.specfialCol].Value));
                int tempCol = Court.excel.specfialCol + 1;
                if (match.Success)
                {
                    string _string = match.Value;
                    Court.excel.WriteToCell(i, tempCol, match.Value + " %");
                }
                else
                {
                    Court.excel.WriteToCell(i, tempCol, "");
                }
            }

            //remove original cell
            for (int i = Court.excel.WS.UsedRange.Row; i < Court.excel.rows + Court.excel.WS.UsedRange.Row; i++)
            {
                match2 = objNotNumberPattern.Match(Convert.ToString(Court.excel.WS.Cells[i, Court.excel.specfialCol].Value2));
                string[] pattern = new string[] { "：", ':'.ToString(), '*'.ToString(), '.'.ToString(), '"'.ToString() };
                char[] array = string.Join(string.Empty, pattern).ToCharArray();

                //Console.WriteLine(Convert.ToString(match2.Success + " + " + Convert.ToString(Court.excel.WS.Cells[i, Court.excel.specfialCol].Value2))); 

                if (match2.Success)
                {
                    string tempString = Convert.ToString(Court.excel.WS.Cells[i, Court.excel.specfialCol].Value2);
                    tempString = tempString.Replace("%", "").Replace(match2.Value, "").Trim(array);

                    Court.excel.WriteToCell(i, Court.excel.specfialCol, tempString);
                }
                Excel.processedRows += 1;
                UpdateProcessedRow();
            }
        }
    }
}
