using Microsoft.Win32;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
namespace Group4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Akhmedova.xaml
    /// </summary>
    public partial class _4337_Akhmedova : Window
    {
        public _4337_Akhmedova()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;

            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ServiceEntities serviceEntities = new ServiceEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    int id = int.Parse(list[i, 0]);
                    int price = int.Parse(list[i, 3]);
                    string category = "";
                    if (price >= 0 && price <= 350)
                    {
                        category = "Категория 1";
                    }
                    else if (price >= 250 && price <= 800)
                    {
                        category = "Категория 2";
                    }
                    else if (price > 800)
                    {
                        category = "Категория 3";
                    }
                    else
                    {
                        category = "неправильно.";
                    }
                    serviceEntities.Service.Add(new Service()
                    {
                        ID = id,
                        Name_service = list[i, 1],
                        Type_service = list[i, 2],
                        Price = list[i, 3],
                        Category = category
                    });
                }
                serviceEntities.SaveChanges();
                MessageBox.Show($"Молодец!{serviceEntities.Service.Count()}");
            }
        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            using (ServiceEntities serviceEntities = new ServiceEntities())
            {
                var category = serviceEntities.Service.GroupBy(x => x.Category).ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = category.Count;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < category.Count; i++)
                {
                    var categori = category[i];
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = categori.Key.Length > 31 ? categori.Key.Substring(0, 31) : categori.Key;

                    worksheet.Cells[1, 1] = "ID";
                    worksheet.Cells[1, 2] = "Название услуги";
                    worksheet.Cells[1, 3] = "Вид услуги";
                    worksheet.Cells[1, 4] = "Стоимость";
                    worksheet.Cells[1, 5] = "Категории";

                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];
                    headerRange.Font.Bold = true;

                    int row = 2;

                    foreach (var service in categori)
                    {
                        worksheet.Cells[row, 1] = service.ID;
                        worksheet.Cells[row, 2] = service.Name_service;
                        worksheet.Cells[row, 3] = service.Type_service;
                        worksheet.Cells[row, 4] = service.Price;
                        worksheet.Cells[row, 5] = service.Category;
                        row++;
                    }

                    worksheet.Cells[row, 1].FormulaLocal = $"=СЧЁТ(D2:D{row - 1})";
                    worksheet.Cells[row, 1].Font.Bold = true;

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row - 1, 5]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Columns.AutoFit();
                }

                app.Visible = true;
            }
        }
    }
}
