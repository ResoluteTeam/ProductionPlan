﻿using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionPlan
{
    public partial class Form1 : Form
    { 
        Excel.Application app;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;

        int sheetscount;
        int lastRow, lastColumn;

        int products;
        int orders;
        int operations;

        public Form1()
        {
            this.DoubleBuffered = true;
            InitializeComponent();

            products = Convert.ToInt32(textBox1.Text);
            operations = Convert.ToInt32(textBox2.Text);
            orders = Convert.ToInt32(textBox3.Text);

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            updateDataGrid();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
/*          workbook.Close(false, false, false);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
            workbook = null;
            worksheet = null;
            System.GC.Collect();
*/
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (this.tabControl1.SelectedIndex)
            {
                case (0):
                    {
                        Console.Write("000000000000000");
                    };
                    break;
                case (1):
                    {
                        Console.Write("111111111111111");
                    };
                    break;
                case (2):
                    {
                        Console.Write("222222222222222");
                    };
                    break;
            }
        }

        private void changeSheet(int num)
        {
            if (num < sheetscount && num > 0)
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(num);
                lastRow = worksheet.UsedRange.Rows.Count;
                lastColumn = worksheet.UsedRange.Columns.Count;
            }
            else
            {
                Console.Write("Sheets don't exist!!!");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                products = Convert.ToInt32(textBox1.Text);
                updateDataGrid();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox2.Text))
            {
                operations = Convert.ToInt32(textBox2.Text);
                updateDataGrid();
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox3.Text))
            {
                orders = Convert.ToInt32(textBox3.Text);
                updateDataGrid();
            }
        }

        private void updateDataGrid ()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();

            dataGridView4.Rows.Clear();
            dataGridView4.Columns.Clear();

            dataGridView1.ColumnCount = products;
            dataGridView2.ColumnCount = products;

            dataGridView3.ColumnCount = 2;
            dataGridView4.ColumnCount = 2;

            dataGridView3.Columns[0].Name = "Заказ";
            dataGridView4.Columns[0].Name = "Заказ";

            dataGridView3.Columns[1].Name = "Срок";
            dataGridView4.Columns[1].Name = "Приоритетность";

            for (int i = 0; i < products; i++)
            {
                dataGridView1.Columns[i].Name = "Изделие№" + (i + 1).ToString();
            }

            for (int i = 0; i < operations; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].HeaderCell.Value = "Операция№" + (i + 1).ToString();
            }

            for (int i = 0; i < products; i++)
            {
                dataGridView2.Columns[i].Name = "Изделие№" + (i + 1).ToString();
            }

            for (int i = 0; i < orders; i++)
            {
                dataGridView2.Rows.Add();
                dataGridView2.Rows[i].HeaderCell.Value = "Заказ№" + (i + 1).ToString();
            }

            for (int i = 0; i < orders; i++)
            {
                dataGridView3.Rows.Add();
                dataGridView3.Rows[i].HeaderCell.Value = "Заказ№" + (i + 1).ToString();

                dataGridView4.Rows.Add();
                dataGridView4.Rows[i].HeaderCell.Value = "Заказ№" + (i + 1).ToString();
            }
        }

        private void fileFromExcel ()
        {
            app = new Excel.Application();

            try
            {
                workbook = app.Workbooks.Open(Application.StartupPath + @"\ProductionPlan.xlsx");
            }
            catch
            {
                Console.Write("Cannot open ProductionPlan.xlsx");
            }
            sheetscount = workbook.Sheets.Count;
            changeSheet(1);

            products = lastColumn;
            operations = lastRow;

            changeSheet(2);
            orders = lastRow;
        }
    }
}
