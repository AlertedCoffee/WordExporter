using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace Lab3.WordExporter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            commodityViewModel = new CommodityRepository();
            this.DataContext = commodityViewModel;

            commodityViewModel.Commodity.Add(new Commodity() { Id = 1, Count = 3, Price = 10, Product = "Продукт 1"});
            commodityViewModel.Commodity.Add(new Commodity() { Id = 1, Count = 5, Price = 9, Product = "Продукт 2"});
        }

        CommodityRepository commodityViewModel;

        public void SetParam(ref Word.Paragraph paragraph)
        {
            paragraph.Range.Font.Name = "Times new roman";
            paragraph.Range.Font.Size = 14;
            paragraph.Range.InsertParagraphAfter();
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            // создаем приложение ворд
            Word.Application winword = new Word.Application();
            //winword.Visible = true;

            // добавляем документ
            Word.Document document = winword.Documents.Add();

            // добавляем параграф с номером накладной и выбранной датой
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            DateTime? selectDate = DateTime.Now;
            string invoiceNumber = "12";
            if (selectDate != null)
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            else
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber);
            SetParam(ref invoicePar);

            // добавляем параграф с поставщиком
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", SellerTextBox.Text);

            SetParam(ref providerPar);


            // добавляем параграф с потребителем
            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            customerPar.Range.Text = "Покупатель: " + BuyerTextBox.Text;
            SetParam(ref customerPar);


            var Shop = commodityViewModel.Commodity;

            int nRows = Shop.Count;
            Word.Table myTable = document.Tables.Add(customerPar.Range, nRows, 4);
            myTable.Borders.Enable = 1;
            // добавляем данные из таблицы в ворд
            for (int i = 1; i < Shop.Count + 1; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = Shop[i - 1].Id.ToString();
                dataRow[2].Range.Text = Shop[i - 1].Product;
                dataRow[3].Range.Text = Shop[i - 1].Count.ToString();
                dataRow[4].Range.Text = Shop[i - 1].Price.ToString();
            }

            Word.Paragraph totalPricePar = document.Content.Paragraphs.Add();
            totalPricePar.Range.Text = "Итого:      " + TotalLabel.Content.ToString();
            totalPricePar.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            SetParam(ref totalPricePar);


            // указываем в какой файл сохранить
            // TODO - добавьте возможность выбора названия файла
            // и места где его нужно сохранить
            object filename = @"D:\wordExample.doc";
            document.SaveAs(filename);
            document.Close();
            winword.Quit();

            MessageBox.Show("Готово");

        }

        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            TotalLabel.Content = commodityViewModel.Sum;

        }
    }
}
