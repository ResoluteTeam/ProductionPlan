﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;


namespace ProductionPlan
{
    public partial class Form1 : Form
    {
        

        Microsoft.Office.Interop.Excel.Application app;
        Microsoft.Office.Interop.Excel.Workbook workbook;
        Microsoft.Office.Interop.Excel.Worksheet worksheet;

        public Form1()
        {
            this.DoubleBuffered = true;
            InitializeComponent();
         
            app = new Microsoft.Office.Interop.Excel.Application();
            try {
                workbook = app.Workbooks.Open(Application.StartupPath + @"\ProductionPlan.xlsx");
            }
            catch {
                Console.Write("Cannot open ProductionPlan.xlsx");
            }
            worksheet = workbook.ActiveSheet;
            setFirstTable();

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        ~Form1()
        {
                
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            app.ActiveWorkbook.Close(true);
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

        private void setFirstTable()
        {
            int rcount = worksheet.UsedRange.Rows.Count;
            Console.Write("rows - " + rcount + "\n");
            Console.Write("cols - " + worksheet.UsedRange.Columns.Count + "\n");

            dataGridView1.ColumnCount = worksheet.UsedRange.Columns.Count;
            dataGridView2.ColumnCount = worksheet.UsedRange.Columns.Count;
            dataGridView3.ColumnCount = worksheet.UsedRange.Columns.Count;
            dataGridView4.ColumnCount = worksheet.UsedRange.Columns.Count;


            for (int i = 1; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i-1].Name = "Изделие №"+(i).ToString();
                dataGridView1.Rows.Add();
            }
            
            for(int i = 1; i < rcount; i++)
            {

            }

            for (int i = 1; i < dataGridView2.ColumnCount; i++)
            {
                
                dataGridView2.Columns[i - 1].Name = "Изделие №" + (i).ToString();
                dataGridView2.Rows.Add();
            }
            dataGridView3.Columns[0].Name = "Изделие";
            dataGridView3.Columns[1].Name = "Срок";
            for (int i = 1; i < dataGridView3.ColumnCount; i++)
            {
                dataGridView3.Rows.Add();
            }
            dataGridView4.Columns[0].Name = "Изделие";
            dataGridView4.Columns[1].Name = "Приоритетность";
            for (int i = 1; i < dataGridView4.ColumnCount; i++)
            {
                dataGridView4.Rows.Add();
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
