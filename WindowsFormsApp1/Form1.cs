using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string invoiceNumber;
        List<Commodity> Shop;
        double totalSum;
        public Form1()
        {
            InitializeComponent();
            invoiceNumber = (new Random().Next(1, 601)).ToString();
            label4.Text = $"Заказ № {invoiceNumber}";
            Shop = new List<Commodity>();
        }

        public void CreateWordDocument()
        {
            // создаем приложение ворд
            Word.Application winword = new Word.Application();
            //winword.Visible = true;

            // добавляем документ
            Word.Document document = winword.Documents.Add();

            // добавляем параграф с номером накладной и выбранной датой
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            DateTime? selectDate = DateTime.Now;

            if (selectDate != null)
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            else
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber);
            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            // добавляем параграф с поставщиком
            string PurchasertxtBox = textBox1.Text;
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", PurchasertxtBox);
            providerPar.Range.Font.Name = "Times new roman";
            providerPar.Range.Font.Size = 14;
            providerPar.Range.InsertParagraphAfter();

            // добавляем параграф с потребителем
            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            string ProvidertxtBox = textBox2.Text;
            customerPar.Range.Text = "Покупатель: " + ProvidertxtBox;
            customerPar.Range.Font.Name = "Times new roman";
            customerPar.Range.Font.Size = 14;
            customerPar.Range.InsertParagraphAfter();

            // формируем таблицу
            // количество колонок - 5
            // количество строк - nRows
            List<Commodity> Shop = new List<Commodity>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    // Извлечение данных из DataGridView и создание объекта Commodity.
                    Commodity commodity = new Commodity
                    {
                        Id = Convert.ToInt32(row.Cells["Id"].Value),
                        Product = Convert.ToString(row.Cells["Product"].Value),
                        Count = Convert.ToInt32(row.Cells["Count"].Value),
                        Price = Convert.ToDouble(row.Cells["Price"].Value),
                        Sum = Convert.ToInt32(row.Cells["Count"].Value) * Convert.ToDouble(row.Cells["Price"].Value)
                    };
                    Shop.Add(commodity);
                }
            }

            int nRows = Shop.Count;
            Word.Table myTable = document.Tables.Add(customerPar.Range, nRows, 5);
            myTable.Borders.Enable = 1;
            // добавляем данные из таблицы в ворд
            for (int i = 1; i < Shop.Count + 1; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = Shop[i - 1].Id.ToString();
                dataRow[2].Range.Text = Shop[i - 1].Product;
                dataRow[3].Range.Text = Shop[i - 1].Count.ToString();
                dataRow[4].Range.Text = Shop[i - 1].Price.ToString();
                dataRow[5].Range.Text = Shop[i-1].Sum.ToString();
            }
            Word.Paragraph totalSum = document.Content.Paragraphs.Add();
            string result = label3.Text;
            totalSum.Range.Text = result;
            totalSum.Range.Font.Name = "Times new roman";
            totalSum.Range.Font.Size = 14;
            totalSum.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            totalSum.Range.InsertParagraphAfter();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word файлы (*.doc)|*.doc|Все файлы (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                document.SaveAs(saveFileDialog.FileName);
                document.Close();
                winword.Quit();
            }
            else
            {
                MessageBox.Show("Error");
            }
        }


        private void CreateExcelDocument()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Добавляем заголовок
            worksheet.Cells[1, 1] = "Расходная накладная № " + invoiceNumber;
            worksheet.Cells[1, 1].Font.Name = "Times new roman";
            worksheet.Cells[1, 1].Font.Size = 14;

            // Добавляем поставщика
            worksheet.Cells[2, 1] = "Поставщик: " + textBox1.Text;
            worksheet.Cells[2, 1].Font.Name = "Times new roman";
            worksheet.Cells[2, 1].Font.Size = 14;

            // Добавляем потребителя
            worksheet.Cells[3, 1] = "Покупатель: " + textBox2.Text;
            worksheet.Cells[3, 1].Font.Name = "Times new roman";
            worksheet.Cells[3, 1].Font.Size = 14;

            // Добавляем заголовки столбцов
            worksheet.Cells[4, 1] = "Id";
            worksheet.Cells[4, 2] = "Product";
            worksheet.Cells[4, 3] = "Count";
            worksheet.Cells[4, 4] = "Price";
            worksheet.Cells[4, 5] = "Sum";

            // Добавляем данные из таблицы
            int row = 5; // Начинаем с пятой строки, так как четвертая строка используется для заголовков
            foreach (DataGridViewRow dataGridViewRow in dataGridView1.Rows)
            {
                if (!dataGridViewRow.IsNewRow)
                {
                    worksheet.Cells[row, 1] = dataGridViewRow.Cells["Id"].Value;
                    worksheet.Cells[row, 2] = dataGridViewRow.Cells["Product"].Value;
                    worksheet.Cells[row, 3] = dataGridViewRow.Cells["Count"].Value;
                    worksheet.Cells[row, 4] = dataGridViewRow.Cells["Price"].Value;
                    worksheet.Cells[row, 5] = dataGridViewRow.Cells["Sum"].Value;

                    row++;
                }
            }

            // Добавляем итоговую сумму
            string totalSum = label3.Text;
            worksheet.Cells[row + 1, 1] = totalSum;
            worksheet.Cells[row + 1, 1].Font.Name = "Times new roman";
            worksheet.Cells[row + 1, 1].Font.Size = 14;
            worksheet.Cells[row + 1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            // Сохраняем файл
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Error");
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            CreateWordDocument();
            totalSum = 0;
            dataGridView1.Rows.Clear();
            label3.Text = $"Итоговая сумма: {totalSum}";
            textBox1.Text = "";
            textBox2.Text = "";
            invoiceNumber = (new Random().Next(1, 601)).ToString();
            label4.Text = $"Заказ № {invoiceNumber}";
            Shop.Clear();
        }


        private void RecalculateTotalSum()
        {
            totalSum = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    int count;
                    double price;

                    if (int.TryParse(Convert.ToString(row.Cells["Count"].Value), out count) &&
                        double.TryParse(Convert.ToString(row.Cells["Price"].Value), out price))
                    {
                        double sum = count * price;
                        row.Cells["Sum"].Value = sum;
                        totalSum += sum;
                    }
                }
            }
            label3.Text = $"Итоговая сумма: {totalSum}";
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            RecalculateTotalSum();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (dataGridView1.Columns[e.ColumnIndex].Name == "Price" || dataGridView1.Columns[e.ColumnIndex].Name == "Count")
                {
                    RecalculateTotalSum();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateExcelDocument();
            totalSum = 0;
            dataGridView1.Rows.Clear();
            label3.Text = $"Итоговая сумма: {totalSum}";
            textBox1.Text = "";
            textBox2.Text = "";
            invoiceNumber = (new Random().Next(1, 601)).ToString();
            label4.Text = $"Заказ № {invoiceNumber}";
            Shop.Clear();
        }
    }


}
