using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTransform
{
    public partial class Form1 : Form
    {
        private System.Drawing.Point _start_point = new System.Drawing.Point(0, 0);
        private bool _mouseDown;

        public Form1()
        {
            InitializeComponent();
            panel3.Visible = false;
            Court.LABEL = label5;
            UpdateExcel.InitializeComponent();
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (Court.patternSelected)
            {
                String[] _dropped_files = (String[])e.Data.GetData(DataFormats.FileDrop);

                foreach (String _eachFile in _dropped_files)
                {
                    string _file_name = getFileName(_eachFile);
                    listBox2.Items.Add(_file_name);
                    transformDate(Path.GetFullPath(_file_name), Path.GetExtension(_file_name));
                }
            }
            else
            {
                MessageBox.Show("Please select a pattern first.");
            }
            
        }

        private void transformDate(string path, string ext)
        {
            if (string.Equals(ext, ".xlsx", StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(ext, ".xls", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    Court.excel = new Excel(@path, 1);
                    String[,] _temp = Court.excel.ReadAll();

                    if (Court.numberSeparated)
                    {
                        for (int j = 0; j < Court.excel.cols; j++)
                        {
                            string _expression = @"^[\p{L}]+(.*)\d+[.]?\d*%?$";
                            Regex _objNotNumberPattern = new Regex(_expression);

                            if (_objNotNumberPattern.IsMatch(_temp[1, j]))
                            {
                                Court.excel.specfialCol = j + 1;
                                break;
                            }
                        }
                        Court.excel.pushCol(Court.excel.specfialCol);
                        UpdateExcel.UpdateSpecialReplace();

                        listBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(path) + "...processed successfully!");
                    }
                    else if (Court.lineSeparated)
                    {

                        //Court.excel.specfialCol = UpdateExcel.UpdateSpecialColumn();
                        //if (Court.excel.specfialCol == 0)
                        //{
                        //    MessageBox.Show("Coundn't find a column with multiple lines.");
                        //}
                        //Court.excel.pushCol(Court.excel.specfialCol);
                        //UpdateExcel.UpdateSpecialSeparate();
                        UpdateExcel.UpdateSpecialSeparateRandom();
                    }

                    Court.excel.SaveAs(path.Insert(path.LastIndexOf("."), "_transformed"));  
                }
                catch (Exception)
                {
                    throw;
                }
                finally {
                    Court.excel.Close();
                }
            }
            else
            {
                MessageBox.Show("Some files are not in .xlsx format.");
                listBox2.Items.RemoveAt(listBox2.Items.Count - 1);
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private string getFileName(string path)
        {
            return System.IO.Path.GetFullPath(path);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This program is designed for GYQ.\n\nPlease contact her if any corrections need to be made.Thank you for using.\n\nAuthor: GC\nContact:GYQ");
            panelLeft.Height = button1.Height;
            panelLeft.Top = button1.Top;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }


        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (_mouseDown)
            {
                System.Drawing.Point P = PointToScreen(e.Location);
                Location = new System.Drawing.Point(P.X - _start_point.X, P.Y - _start_point.Y);
            }
        }

        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            _mouseDown = false;
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            _mouseDown = true;
            _start_point = new System.Drawing.Point(e.X, e.Y);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Court.excel == null) {
                listBox2.Items.Clear();
                listBox1.Items.Clear();
                UpdateExcel.UpdateProcessedRow();
                MessageBox.Show("You haven't started processing any file(s) yet.");
            }
            else
            {
                listBox2.Items.Clear();
                listBox1.Items.Clear();
                Court.excel.Clear();
                UpdateExcel.UpdateProcessedRow();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (panel3.Visible) panel3.Visible = false;
            else panel3.Visible = true;
        }


        private void Form1_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Court.currentExtraction = Court.expressionNumberExtraction;
            Court.currentExtractionSymbol = Court.expressionNumberExtractionSymbol;
            Court.patternSelected = true;
            Court.numberSeparated = true;
            Court.lineSeparated = false;
            panel3.Visible = false;
        }

        private void panel3_MouseEnter(object sender, EventArgs e)
        {
            panel3.Visible = true;
        }

        private void panel3_MouseLeave(object sender, EventArgs e)
        {
            panel3.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Court.patternSelected = true;
            Court.lineSeparated = true;
            Court.numberSeparated = false;
            panel3.Visible = false;
        }
    }
}
