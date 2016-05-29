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

        int[] terms;

        int sheetscount;
        int lastRow, lastColumn;

        int date;
        int products = 2;
        int orders = 2;
        int operations = 2;
        int maxTime;

        public Form1()
        {
            this.DoubleBuffered = true;
            InitializeComponent();

            products = Convert.ToInt32(textBox1.Text);
            operations = Convert.ToInt32(textBox2.Text);
            orders = Convert.ToInt32(textBox3.Text);
            terms = new int[orders];

            date = 0;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            updateDataGrid();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex++;
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
                if (Convert.ToInt32(textBox1.Text) <= 250)
                {
                    products = Convert.ToInt32(textBox1.Text);
                    updateDataGrid();
                } else
                {
                    MessageBox.Show("Некорректный ввод!\nВведите число от 0 до 250");
                    textBox1.Text = "2";
                }
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox2.Text) && Convert.ToInt32(textBox2.Text) != 0)
            {
                if (Convert.ToInt32(textBox2.Text) <= 250)
                {
                    operations = Convert.ToInt32(textBox2.Text);
                    updateDataGrid();
                } else
                {
                    MessageBox.Show("Некорректный ввод!\nВведите число от 0 до 250");
                    textBox2.Text = "2";
                }
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox3.Text) && Convert.ToInt32(textBox3.Text) != 0)
            {
                if (Convert.ToInt32(textBox3.Text) <= 250)
                {
                    orders = Convert.ToInt32(textBox3.Text);
                    terms = new int[orders];
                    updateDataGrid();
                } else
                {
                    MessageBox.Show("Некорректный ввод!\nВведите число от 0 до 250");
                    textBox3.Text = "2";
                }
            }
        }

        private void updateDataGrid()
        {
            dataGridView1.DefaultCellStyle.NullValue = "0";
            dataGridView2.DefaultCellStyle.NullValue = "0";
            dataGridView3.DefaultCellStyle.NullValue = "0";
            dataGridView4.DefaultCellStyle.NullValue = "0";

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

            terms = new int[orders];
            getDataFromProductGrid();
            getDataFromOrdersGrid();
            
        }

        private void fileFromExcel(string location)
        {
            app = new Excel.Application();

            try
            {
                workbook = app.Workbooks.Open(location);
            }
            catch
            {
                Console.Write("Cannot open ProductionPlan.xlsx");
            }
            sheetscount = workbook.Sheets.Count;
            changeSheet(1);

            products = lastColumn;
            textBox1.Text = products.ToString();
            operations = lastRow;
            textBox2.Text = operations.ToString();

            changeSheet(2);
            orders = lastRow;
            textBox3.Text = orders.ToString();

            updateDataGrid();


            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            for (int i = 0; i < worksheet.UsedRange.Rows.Count; i++)
            {
                for (int j = 0; j < worksheet.UsedRange.Columns.Count; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = worksheet.Cells[i + 1, j + 1].Value.ToString();
                }
            }

            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(2);
            for (int i = 0; i < worksheet.UsedRange.Rows.Count; i++)
            {
                for (int j = 0; j < worksheet.UsedRange.Columns.Count; j++)
                {
                    dataGridView2.Rows[i].Cells[j].Value = worksheet.Cells[i + 1, j + 1].Value.ToString();
                }
            }

            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(3);

            for (int i = 0; i < orders; i++)
            {
                dataGridView3.Rows[i].Cells[0].Value = worksheet.Cells[i + 1, 1].Value.ToString();
                dataGridView4.Rows[i].Cells[0].Value = worksheet.Cells[i + 1, 2].Value.ToString();
            }
            workbook.Close();
            app.Quit();
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

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                Product tempProduct = new Product(operations);
                int[] duration = new int[operations];

                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    duration[j] = Convert.ToInt32(dataGridView1.Rows[j].Cells[i].Value); 
                }

                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    tempProduct.Duration.SetValue(duration[j], j);
                }

                productList.Add(tempProduct);
            }
        }
        private void getDataFromOrdersGrid()
        {
            date = 0;
            if (terms[0] == 0)
            {
                for (int i = 0; i < orders; i++)
                {
                    terms[i] = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);

                }
            }
            ordersList = new List<Order>();

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                Order tempOrder = new Order(products);
                int[] amount = new int[products];

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

                if (radioButton2.Checked)
                    tempOrder.Time = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);
                else tempOrder.Time = terms[i];

                ordersList.Add(tempOrder);
            }

            int temp = 0;

            if (radioButton2.Checked)
            {
                for (int i = 0; i < dataGridView3.RowCount; i++)
                {
                    if (Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value) > temp)
                    {
                        temp = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);
                    }
                }
            }
            else
            {
                for(int i = 0; i < terms.Count(); i++)
                {
                    if (temp < terms[i])
                        temp = terms[i];
                }
            }
            maxTime = temp;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                getDataFromProductGrid();
            }

            if (tabControl1.SelectedIndex == 2)
            {
                getDataFromOrdersGrid();
                createResultGrid();
                if (radioButton1.Checked)
                    calculateByPriority();
                if (radioButton2.Checked)
                    calculateByTime();
            }
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            int newInteger;

            if (!int.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0)
            {
                e.Cancel = true;
                MessageBox.Show("Некорректный ввод!");
            }
        }
        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            int newInteger;

            if (!int.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0)
            {
                e.Cancel = true;
                MessageBox.Show("Некорректный ввод!");
            }
        }
        private void dataGridView3_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            int newInteger;

            if (!int.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0)
            {
                e.Cancel = true;
                MessageBox.Show("Некорректный ввод!");
            }
        }
        private void dataGridView4_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            float newInteger;

            if (!float.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0 || newInteger > 1)
            {
                e.Cancel = true;
                MessageBox.Show("Некорректный ввод!\nВведите число от 0 до 1");
            }

        }

        private void createResultGrid()
        {
            dataGridView5.Rows.Clear();
            dataGridView5.Columns.Clear();

            if (radioButton2.Checked)
            {
                date = 0;
                for (int i = 0; i < ordersList.Count; i++)
                {
                    if (ordersList.ElementAt(i).Enabled)
                        if (date < ordersList.ElementAt(i).Time)
                            date = ordersList.ElementAt(i).Time;
                }
            }
            else date = ordersList.ElementAt(0).Time;
            dataGridView5.ColumnCount = date;
            dataGridView5.RowCount = orders * products * operations + operations + 1;

            int index;
            index = date;

            for (int i = index - 1; i >= 0; i--)
            {
                dataGridView5.Columns[i].Name = "День " + (index - i).ToString();
                dataGridView5.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (int i = 0; i < orders; i++)
            {
                dataGridView5.Rows[i * products * operations].HeaderCell.Value = " Заказ №" + (i + 1).ToString();
                for (int j = 0; j < products; j++)
                {
                    dataGridView5.Rows[i * products * operations + j * operations].HeaderCell.Value += " Изделие №" + (j + 1).ToString();
                    for (int k = 0; k < operations; k++)
                    {
                        dataGridView5.Rows[i * products * operations + j * operations + k].HeaderCell.Value += " Операция №" + (k + 1).ToString();
                    }
                }
            }

            for (int i = 0; i < operations; i++)
            {
                dataGridView5.Rows[orders * products * operations + i + 1].HeaderCell.Value = "Станок №" + (i + 1).ToString();
            }
            dataGridView5.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void calculateByPriority()
        {
            ordersList = ordersList.OrderByDescending(Order => Order.Priority).ToList();
            int currentProduct;
            while (ordersList.Any())
            {
                int currentDate = date - 1;
                currentProduct = -1;
                for (int i = 0; i < products; i++)
                {
                    if (ordersList.ElementAt(0).Products[i] != 0)
                    {
                        currentProduct = i; //Выбираем продукт из списка
                        i = products;
                    }
                }

                if (currentProduct < 0)
                {
                    ordersList.RemoveAt(0); //Если количество в списке каждого продукта - 0, идем до след. заказа
                } else // иначе начинаем записывать текущий продукт в таблицу   
                {
                    for (int i = 0; i < operations; i++)
                    {
                        int times = 0;
                        int temp = 0;
                        for (int j = 0; j < operations; j++) //Подсчёт уже отработаных часов на станке на одном заказе
                        {
                            if(currentDate >= 0)
                                times += Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + j].Cells[currentDate].Value);
                        
                        }

                        for (int k = 0; k < orders * products; k++) //Подсчёт уже отработаных часов на станке за весь день на всех заказах
                        {
                            if (currentDate >= 0)
                            {

                                if (i == 0)
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                                else
                                {
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i - 1].Cells[currentDate].Value);
                                }
                            }
                        }

                        if (temp > times)
                            times = temp;

                        int remainder = 0;
                        if (currentProduct >= 0)
                            remainder = productList.ElementAt(currentProduct).Duration.ElementAt(i); //Сколько нужно времени для данной операции

                        while (remainder > 8 - times)
                        {
                            currentDate--;

                            if (currentDate == -1)
                            {
                                date++;
                                for (int n = 0; n < dataGridView3.RowCount; n++)
                                {
                                    terms[n] += 1; //добавление + 1 дня к плану
                                }
                                getDataFromOrdersGrid();
                                ordersList = ordersList.OrderByDescending(Order => Order.Priority).ToList();
                                createResultGrid();
                                currentProduct = -1;
                                i = operations;
                                break;
                            }

                            if (times < 8)
                            {
                                dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate + 1].Value = 
                                    Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate + 1].Value) + (8 - times);
                                remainder = remainder - (8 - times);
                            }
                            
                            times = 0;
                            temp = 0;
                            for (int j = 0; j < operations; j++)
                            {
                                if (currentDate >= 0)
                                    times += Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * products + j].Cells[currentDate].Value);
                            }

                            for (int k = 0; k < orders * products; k++)
                            {
                                if (currentDate >= 0)
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                            }

                            if (temp > times)
                                times = temp;
                        }
                        if (currentDate != -1)
                        {
                            dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate].Value =
                            Convert.ToInt32(dataGridView5.Rows[ordersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate].Value) +
                            remainder;
                        }
                    }

                    if(currentProduct != -1)
                        ordersList.ElementAt(0).Products[currentProduct]--;
                }       
            }
            calculateTime();
        }

        private void calculateByTime()
        {
            ordersList = ordersList.OrderByDescending(Order => Order.Priority).ToList();
            List<Order> tempOrdersList = new List<Order>();
            for (int i = 0; i < ordersList.Count; i++)
            {
                if (ordersList.ElementAt(i).Enabled == true)
                    tempOrdersList.Add(ordersList.ElementAt(i));
            }

            int currentProduct;

            while (tempOrdersList.Any())
            {
                int currentDate = date - 1;
                currentProduct = -1;
                for (int i = 0; i < products; i++)
                {
                    if (tempOrdersList.ElementAt(0).Products[i] != 0)
                    {
                        currentProduct = i; //Выбираем продукт из списка
                        i = products;
                    }
                }

                if (currentProduct < 0)
                {
                    tempOrdersList.RemoveAt(0); //Если количество в списке каждого продукта - 0, идем до след. заказа
                }
                else // иначе начинаем записывать текущий продукт в таблицу   
                {
                    for (int i = 0; i < operations; i++)
                    {
                        int times = 0;
                        int temp = 0;
                        for (int j = 0; j < operations; j++)
                        {
                            if (currentDate >= 0)
                                times += Convert.ToInt32(dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * operations + j].Cells[currentDate].Value);
                        }

                        for (int k = 0; k < orders * products; k++)
                        {
                            if (currentDate >= 0)
                            {
                                if (i == 0)
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                                else
                                {
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i - 1].Cells[currentDate].Value);
                                }
                            }
                        }

                        if (temp > times)
                            times = temp;

                        int remainder = productList.ElementAt(currentProduct).Duration.ElementAt(i);

                        while (remainder > 8 - times)
                        {
                            currentDate--;

                            if (currentDate == -1)
                            {
                                for (int n = 0; n < ordersList.Count; n++)
                                {
                                    if (ordersList.ElementAt(n).Index == tempOrdersList.ElementAt(0).Index)
                                    {
                                        ordersList.ElementAt(n).Enabled = false;
                                    }

                                    int[] amount = new int[products];
                                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                                    {
                                        amount[j] = Convert.ToInt32(dataGridView2.Rows[ordersList.ElementAt(n).Index].Cells[j].Value);
                                    }
                                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                                    {
                                        ordersList.ElementAt(n).Products.SetValue(amount[j], j);
                                    }
                                }

                                currentProduct = -1;
                                i = operations;

                                tempOrdersList.Clear();

                                for (int n = 0; n < ordersList.Count; n++)
                                {
                                    if (ordersList.ElementAt(n).Enabled == true)
                                        tempOrdersList.Add(ordersList.ElementAt(n));
                                }

                                tempOrdersList = tempOrdersList.OrderByDescending(Order => Order.Priority).ToList();

                                int tempTime = 0;
                                for (int n = 0; n < tempOrdersList.Count; n++)
                                {
                                    if (tempOrdersList.ElementAt(n).Enabled)
                                        if (tempOrdersList.ElementAt(n).Time > tempTime)
                                            tempTime = tempOrdersList.ElementAt(n).Time;
                                }
                                date = tempTime;
                                createResultGrid();
                                break;
                            }

                            if (times < 8)
                            {
                                dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate + 1].Value =
                                    Convert.ToInt32(dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate + 1].Value) + (8 - times);
                                remainder = remainder - (8 - times);
                            }

                            times = 0;
                            temp = 0;
                            for (int j = 0; j < operations; j++)
                            {
                                if (currentDate >= 0)
                                    times += Convert.ToInt32(dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * products + j].Cells[currentDate].Value);
                            }

                            for (int k = 0; k < orders * products; k++)
                            {
                                if (currentDate >= 0)
                                    temp += Convert.ToInt32(dataGridView5.Rows[k * operations + i].Cells[currentDate].Value);
                            }

                            if (temp > times)
                                times = temp;
                        }
                        if (currentDate != -1)
                        {
                            dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate].Value =
                            Convert.ToInt32(dataGridView5.Rows[tempOrdersList.ElementAt(0).Index * products * operations + currentProduct * operations + i].Cells[currentDate].Value) +
                            remainder;
                        }
                    }

                    if (currentProduct != -1)
                        tempOrdersList.ElementAt(0).Products[currentProduct]--;
                }
            }
            calculateTime();
        }

        private void calculateTime()
        {
            int day = 0;
            int j = 0;
            int i = 0;
            int k = 0;
            int index;

            index = date;

            while (day < index)
            { 
                while (i < operations)
                {
                    while (j < orders)
                    {
                        while (k < products)
                        {
                            dataGridView5.Rows[orders * products * operations + i + 1].Cells[day].Value =
                                (Convert.ToInt32(dataGridView5.Rows[orders * products * operations + i + 1].Cells[day].Value)
                                + Convert.ToInt32(dataGridView5.Rows[j * products * operations + k * operations + i].Cells[day].Value)).ToString();
                            k++;
                        }
                        j++;
                        k = 0;
                    }
                    i++;
                    j = 0;
                }
                i = 0;
                day++;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                getDataFromProductGrid();
                getDataFromOrdersGrid();

                createResultGrid();
                calculateByTime();
            }
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            terms = new int[orders];
            for (int i = 0; i < dataGridView3.RowCount; i++)
            {
                terms[i] = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);

            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            terms = new int[orders];
            for (int i = 0; i < dataGridView3.RowCount; i++)
            {
                terms[i] = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);

            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileFromExcel(openFileDialog1.FileName);
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application saveApp;
                Excel.Workbook saveWorkbook;
                Excel.Worksheet saveWorksheet;

                saveApp = new Excel.Application();
                saveApp.DisplayAlerts = false;

                object misValue = System.Reflection.Missing.Value;

                saveWorkbook = saveApp.Workbooks.Add(System.Reflection.Missing.Value);
                saveWorksheet = (Excel.Worksheet)saveWorkbook.Worksheets.get_Item(1);

                for (int i = 0; i < dataGridView5.ColumnCount; i++) {
                    saveWorksheet.Cells[1, i + 2].Value = dataGridView5.Columns[i].HeaderCell.Value;
                }
                for (int i = 0; i < dataGridView5.RowCount; i++)
                {
                    saveWorksheet.Cells[i + 2, 1].Value = dataGridView5.Rows[i].HeaderCell.Value;
                }

                for (int i = 0; i < dataGridView5.RowCount; i ++)
                {
                    for (int j = 0; j < dataGridView5.ColumnCount; j++)
                    {
                        saveWorksheet.Cells[i + 2, j + 2].Value = dataGridView5.Rows[i].Cells[j].Value;
                    }
                }
                saveWorksheet.Columns.AutoFit();
                saveWorkbook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                saveWorkbook.Close();
                saveApp.Quit();
                
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                getDataFromProductGrid();
                getDataFromOrdersGrid();

                createResultGrid();
                calculateByPriority();
            }
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

        bool enabled = true;

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

        public bool Enabled
        {
            get
            {
                return enabled;
            }

            set
            {
                enabled = value;
            }
        }

        public Order(int prodCount)
        {
            products = new int[prodCount];
        }
    }
}
