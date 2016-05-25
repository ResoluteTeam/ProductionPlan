using System;
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

        public Form1()
        {
            this.DoubleBuffered = true;
            InitializeComponent();
         
            app = new Excel.Application();

            try {
                workbook = app.Workbooks.Open(Application.StartupPath + @"\ProductionPlan.xlsx");
            }
            catch {
                Console.Write("Cannot open ProductionPlan.xlsx");
            }
            sheetscount = workbook.Sheets.Count;
            changeSheet(1);
            setFirstTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            workbook.Close(false, false, false);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
            workbook = null;
            worksheet = null;
            System.GC.Collect();

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
            }

            for (int i = 0; i < rcount; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].HeaderCell.Value = "Операция " + (i + 1).ToString();
            }

            for (int i = 1; i < dataGridView2.ColumnCount; i++)
            {
                
                dataGridView2.Columns[i - 1].Name = "Изделие №" + (i).ToString();
                dataGridView2.Rows.Add();
            }
            dataGridView3.Columns[0].Name = "Заказ";
            dataGridView3.Columns[1].Name = "Срок";
            for (int i = 1; i < dataGridView3.ColumnCount; i++)
            {
                dataGridView3.Rows.Add();
            }
            dataGridView4.Columns[0].Name = "Заказ";
            dataGridView4.Columns[1].Name = "Приоритетность";
            for (int i = 1; i < dataGridView4.ColumnCount; i++)
            {
                dataGridView4.Rows.Add();
            }
        }

        private void changeSheet(int num)
        {
            if (num < sheetscount && num > 0)
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(num);
                int lastRow = worksheet.UsedRange.Rows.Count;
                int lastColumn = worksheet.UsedRange.Columns.Count;
            }
            else
            {
                Console.Write("Sheets don't exist!!!");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
