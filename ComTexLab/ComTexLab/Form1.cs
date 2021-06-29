using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ComTexLab
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            var E = new Excel.Application();
            E.Visible = true;
            E.Workbooks.Open(openFileDialog1.FileName);
            var Sh = E.ActiveSheet;
            Sh.Name = "COM";
            string[] arr = new string[4];
            dataGridView1.ColumnCount = 4;
            dataGridView1.Rows.Clear();
            int r = 0;
            do
            {
                r++;
                arr[0] = Sh.Cells[r + 1, 1].Text;
                if (arr[0] == "") break;
                arr[1] = Sh.Cells[r + 1, 2].Text;
                arr[2] = Sh.Cells[r + 1, 3].Text;
                arr[3] = Sh.Cells[r + 1, 4].Text;
                dataGridView1.Rows.Add(arr[0], arr[1], arr[2], arr[3]);
            } while (true);
            r--; 
            var Ch = E.Charts.Add();
            Ch.Location(Excel.XlChartLocation.xlLocationAsObject, "COM");
            Ch = E.ActiveChart;
            Ch.SeriesCollection(1).Delete();
            Ch.HasTitle = 1;
            Ch.HasLegend = false;
            Ch.ChartTitle.Text = "Успеваемость студентов";
            Ch.Axes(1).HasTitle = true;
            Ch.Axes(1).AxisTitle.Text = "Студенты";
            Ch.Axes(2).HasTitle = true;
            Ch.Axes(2).AxisTitle.Text = "Рейтинг";
            var W = new Word.Application();
            W.Visible = false;
            var D = W.Documents.Add();
            var t = W.Selection;
            t.Font.Bold = -1;
            t.Font.Size = 24;
            t.TypeText("Данные о студентах.");
            t.TypeParagraph();
            t.Font.Bold = 0;
            t.Font.Size = 14;
            for (int k = 1; k <= r; k++)
            {
                t.TypeText(k.ToString() + " : ");
                t.TypeText(Sh.Cells[k + 1, 2].Value.ToString() + " ");
                t.TypeText(Sh.Cells[k + 1, 3].Text + " ");
                t.TypeText(Sh.Cells[k + 1, 4].Value.ToString());
                t.TypeParagraph();
            }
            t.TypeParagraph();
            t.TypeText("Всего " + r.ToString() + " студентов.");
            t.TypeParagraph();
            Ch.ChartArea.Select();
            Ch.ChartArea.Copy();
            t.TypeParagraph();
            t.Paste();
            D.SaveAs(Path.GetDirectoryName(openFileDialog1.FileName) + "\\Test");
            W.Quit(); W = null;
            E.Quit(); E = null;
        }
    }
}
