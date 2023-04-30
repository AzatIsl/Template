using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Office.Interop;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Islamkhuzin.xaml
    /// </summary>
    public partial class _4333_Islamkhuzin : Window
    {
        public _4333_Islamkhuzin()
        {
            InitializeComponent();
        }


        private void Import_Click(object sender, RoutedEventArgs e)
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
            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (GITZd2Entities gitzd2Entities = new GITZd2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    if (i == 0 || string.IsNullOrWhiteSpace(list[i,0]))
                        continue;
                    gitzd2Entities.zad2_table.Add(new zad2_table()
                    {
                        ID = list[i, 0],
                        OrderCode = list[i, 1],
                        DateOfCreation = list[i, 2],
                        OrderTime = list[i, 3],
                        ClientCode = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        ClosingDate = list[i, 7],
                        RentalTime = list[i, 8]
                    });
                }
                gitzd2Entities.SaveChanges();
                MessageBox.Show("Успешный импорт");
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            using (var context = new GITZd2Entities())
            {
                var statuses = context.zad2_table.Select(x => x.Status).Distinct().ToList();
                var app = new Microsoft.Office.Interop.Excel.Application
                {
                    //Отобразить Excel
                    Visible = true,
                    //Количество листов в рабочей книге
                    SheetsInNewWorkbook = statuses.Count()
                };
                Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                int i = 1;
                foreach (var status in statuses)
                {
                    var data = context.zad2_table.Where(x => x.Status == status).ToList().OrderByDescending(x=>Convert.ToInt32(x.ID));

                    Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)app.Worksheets.get_Item(i++);
                    xlWorksheet.Name = status;
                    int lastRow = 0;

                    xlWorksheet.Cells[lastRow + 1, 1] = "ID";
                    xlWorksheet.Cells[lastRow + 1, 2] = "Код заказа";
                    xlWorksheet.Cells[lastRow + 1, 3] = "Дата создания";
                    xlWorksheet.Cells[lastRow + 1, 4] = "Время заказа";
                    xlWorksheet.Cells[lastRow + 1, 5] = "Код клиента";
                    xlWorksheet.Cells[lastRow + 1, 6] = "Услуги";
                    xlWorksheet.Cells[lastRow + 1, 7] = "Статус";
                    xlWorksheet.Cells[lastRow + 1, 8] = "Дата закрытия";
                    xlWorksheet.Cells[lastRow + 1, 9] = "Время проката";
                    lastRow++;

                    foreach (var row in data)
                    {
                        // Добавляем новую строку после последней заполненной строки
                        Microsoft.Office.Interop.Excel.Range range = xlWorksheet.Range["A" + (lastRow + 1).ToString(), "A" + (lastRow + 1).ToString()];
                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                        // Заполняем добавленную строку данными
                        xlWorksheet.Cells[lastRow + 1, 1] = row.ID; 
                        xlWorksheet.Cells[lastRow + 1, 2] = row.OrderCode;
                        xlWorksheet.Cells[lastRow + 1, 3] = row.DateOfCreation;
                        xlWorksheet.Cells[lastRow + 1, 4] = row.OrderTime ;
                        xlWorksheet.Cells[lastRow + 1, 5] = row.ClientCode;
                        xlWorksheet.Cells[lastRow + 1, 6] = row.Services ;
                        xlWorksheet.Cells[lastRow + 1, 7] = row.Status;
                        xlWorksheet.Cells[lastRow + 1, 8] = row.ClosingDate ;
                        xlWorksheet.Cells[lastRow + 1, 9] = row.RentalTime;

                    } 
                }
                // сохранение данных в отдельный файл или лист Excel
                workbook.SaveAs("C:\\Users\\azati\\OneDrive\\Рабочий стол\\Лабораторные работы\\h.xlsx");
                workbook.Close();
                app.Quit();
            }

            return;

        }
    }
}
