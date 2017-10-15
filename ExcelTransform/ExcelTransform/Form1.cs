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
            Court.LABEL = label5;
            Court.LABEL.Text = "Done: " + Court.PROCESSED_ROWS;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            String[] _dropped_files = (String[])e.Data.GetData(DataFormats.FileDrop);

            foreach (String _eachFile in _dropped_files)
            {
                string _file_name = getFileName(_eachFile);
                listBox2.Items.Add(_file_name);
                transformDate(Path.GetFullPath(_file_name), Path.GetExtension(_file_name));
            }
        }

        private void transformDate(string path, string ext)
        {
            if (string.Equals(ext, ".xlsx", StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(ext, ".xls", StringComparison.CurrentCultureIgnoreCase))
            {
                Excel _excel = new Excel(@path, 1);
                String[,] _temp = _excel.ReadAll();

                for (int j = 0; j < _excel.cols; j++)
                {
                    string _expression = @"^[\p{L}]+(.*)\d+[.]?\d*%?$";
                    Regex _objNotNumberPattern = new Regex(_expression);

                    if (_objNotNumberPattern.IsMatch(_temp[1, j]))
                    {
                        _excel.specfialCol = j + 1;
                        break;
                    }
                }

                _excel.pushCol(_excel.specfialCol);
                _excel.specialReplace();
                _excel.SaveAs(path.Insert(path.LastIndexOf("."), "_transformed"));
                _excel.Close();

                listBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(path) + "...processed successfully!");
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
            MessageBox.Show("This program is designed to be used by GYQ.\nPlease contact her if any corrections need to be made.\nThank you for using.");
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
            listBox2.Items.Clear();
            listBox1.Items.Clear();
            Court.PROCESSED_ROWS = 0;
            Court.LABEL.Text = "Done: " + Court.PROCESSED_ROWS;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
