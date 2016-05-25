using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductionPlan
{
    public class Product
    {
        int[] duration;

        public Product(int operationsCount)
        {
            duration = new int[operationsCount];
        }

        public int[] Duration
        {
            get
            {
                return duration;
            }

            set
            {
                duration = value;
            }
        }
    }

    public class Order
    {
        int[] products;
        float priority;

        public float Priority
        {
            get
            {
                return priority;
            }

            set
            {
                priority = value;
            }
        }

        public int[] Products
        {
            get
            {
                return products;
            }

            set
            {
                products = value;
            }
        }

        public Order(int prodCount)
        {
            products = new int[prodCount];
        }
    }

    public partial class Form1 : Form
    {
        Excel.Application app;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;

        List<Order> ordersList;

        int sheetscount;
        int lastRow, lastColumn;

        int products;
        int orders;
        int operations;

        List<List<int>> productToOperations = new List<List<int>>();

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
            getData();
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
            if (!string.IsNullOrWhiteSpace(textBox1.Text) && Convert.ToInt32(textBox1.Text) != 0)
            {
                products = Convert.ToInt32(textBox1.Text);
                updateDataGrid();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox2.Text) && Convert.ToInt32(textBox2.Text) != 0)
            {
                operations = Convert.ToInt32(textBox2.Text);
                updateDataGrid();
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox3.Text) && Convert.ToInt32(textBox3.Text) != 0)
            {
                orders = Convert.ToInt32(textBox3.Text);
                updateDataGrid();
            }
        }

        private void updateDataGrid()
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
            dataGridView3.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView4.Columns[0].Name = "Заказ";
            dataGridView4.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView3.Columns[1].Name = "Срок";
            dataGridView3.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView4.Columns[1].Name = "Приоритетность";
            dataGridView4.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;

            for (int i = 0; i < products; i++)
            {
                dataGridView1.Columns[i].Name = "Изделие№" + (i + 1).ToString();
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (int i = 0; i < operations; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].HeaderCell.Value = "Операция№" + (i + 1).ToString();
            }

            for (int i = 0; i < products; i++)
            {
                dataGridView2.Columns[i].Name = "Изделие№" + (i + 1).ToString();
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
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

            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView3.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView4.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void fileFromExcel()
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

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void getData()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                List<int> temp = new List<int>();
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    int x = Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value);
                    temp.Add(x);
                }
                productToOperations.Add(temp);
            }
        }
    }
}
