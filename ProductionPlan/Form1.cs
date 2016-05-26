using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ProductionPlan
{
    public partial class Form1 : Form
    {
        Excel.Application app;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;

        List<Order> ordersList;
        List<Product> productList;

        int sheetscount;
        int lastRow, lastColumn;

        int products;
        int orders;
        int operations;
        int maxTime;

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
            tabControl1.SelectedIndex++;
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

            dataGridView3.ColumnCount = 1;
            dataGridView4.ColumnCount = 1;

            dataGridView3.Columns[0].Name = "Срок";
            dataGridView3.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView4.Columns[0].Name = "Приоритетность";
            dataGridView4.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

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
            dataGridView5.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            getDataFromOrdersGrid();
            getDataFromProductGrid();
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

        private void getDataFromProductGrid()
        {
            productList = new List<Product>();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                Product tempProduct = new Product(operations);
                int[] duration = new int[operations];

                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    duration[j] = Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value); 
                }

                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    tempProduct.Duration.SetValue(duration[j], j);
                }

                productList.Add(tempProduct);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            getDataFromProductGrid();
            getDataFromOrdersGrid();
            if (tabControl1.SelectedIndex == 2)
            {
                createResultGrid();
                calculateByPriority();   
            }
        }

        private void calculateByPriority()
        {
            ordersList = ordersList.OrderByDescending(Order => Order.Priority).ToList();

            int currentProduct = -1;
            while (ordersList.Any())
            {
                currentProduct = -1;
                for (int i = 0; i < products; i++)
                {
                    if (ordersList.ElementAt(0).Products[i] != 0)
                    {
                        currentProduct = i;
                        i = products;
                    }
                }

                if (currentProduct == -1)
                {
                    ordersList.RemoveAt(0);
                } else
                {
                    for (int i = operations - 1; i >= 0; i--) {
                        int times = 0;
                        for (int j = 0; j < orders * products; j++)
                        {
                            times += Convert.ToInt32(dataGridView5.Rows[i + j * operations].Cells[ordersList.ElementAt(0).Time - 1].Value);
                        }
                        if (Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[ordersList.ElementAt(0).Time - 1].Value) +
                            productList.ElementAt(currentProduct).Duration.ElementAt(i) < 8 - times)
                        {
                            dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[ordersList.ElementAt(0).Time - 1].Value =
                            Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[ordersList.ElementAt(0).Time - 1].Value) +
                            productList.ElementAt(currentProduct).Duration.ElementAt(i);
                        } else
                        {
                           ordersList.ElementAt(0).Time--;
                           dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[ordersList.ElementAt(0).Time - 1].Value =
                           Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[ordersList.ElementAt(0).Time - 1].Value) +
                           productList.ElementAt(currentProduct).Duration.ElementAt(i);
                        }
                    }
                    ordersList.ElementAt(0).Products[currentProduct]--;
                }

                
            }
        }

        private void getDataFromOrdersGrid()
        {
            ordersList = new List<Order>();

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                Order tempOrder = new Order(products);
                int[] amount = new int[products];
                int priority;

                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    amount[j] = Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value);
                }

                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    tempOrder.Products.SetValue(amount[j], j);
                }

                tempOrder.Priority = Convert.ToSingle(dataGridView4.Rows[i].Cells[0].Value);
                tempOrder.Index = i;
                tempOrder.Time = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);

                ordersList.Add(tempOrder);
            }

            int temp = 0;
            for (int i = 0; i < dataGridView3.RowCount; i++)
            {
                if (Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value) > temp)
                {
                    temp = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);
                }
            }
            maxTime = temp;

        }
        private void createResultGrid()
        {
            dataGridView5.Rows.Clear();
            dataGridView5.Columns.Clear();

            dataGridView5.ColumnCount = maxTime;
            dataGridView5.RowCount = orders * products * operations;
            

            for (int i = 0; i < maxTime; i++)
            { 
                dataGridView5.Columns[i].Name = "День " + (i + 1).ToString();
                dataGridView5.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (int i = 0; i < orders; i++)
            {
                dataGridView5.Rows[i * products * operations].HeaderCell.Value = "Заказ№" + (i + 1).ToString();
                for (int j = 0; j < products; j++)
                {
                    dataGridView5.Rows[i * products * operations + j * operations].HeaderCell.Value += " Изделие№" + (j + 1).ToString();
                    for (int k = 0; k < operations; k++)
                    {
                        dataGridView5.Rows[i * products * operations + j * operations + k].HeaderCell.Value += " Операция№" + (k + 1).ToString();
                    }
                }
            }
            dataGridView5.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
    }

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
        int index;
        int time;

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

        public int Index
        {
            get
            {
                return index;
            }

            set
            {
                index = value;
            }
        }

        public int Time
        {
            get
            {
                return time;
            }

            set
            {
                time = value;
            }
        }

        public Order(int prodCount)
        {
            products = new int[prodCount];
        }
    }
}
