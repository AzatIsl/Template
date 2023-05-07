using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Office.Interop;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Text.Json;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Islamkhuzin.xaml
    /// </summary>
    public partial class _4333_Islamkhuzin : System.Windows.Window
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
                    if (i == 0 || string.IsNullOrWhiteSpace(list[i, 0]))
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
                    var data = context.zad2_table.Where(x => x.Status == status).ToList().OrderByDescending(x => Convert.ToInt32(x.ID));

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
                        xlWorksheet.Cells[lastRow + 1, 4] = row.OrderTime;
                        xlWorksheet.Cells[lastRow + 1, 5] = row.ClientCode;
                        xlWorksheet.Cells[lastRow + 1, 6] = row.Services;
                        xlWorksheet.Cells[lastRow + 1, 7] = row.Status;
                        xlWorksheet.Cells[lastRow + 1, 8] = row.ClosingDate;
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

        private void Import_JSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                //DefaultExt = "*.xls;*.xlsx",
                //Filter = "файл Excel",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            var list = JsonSerializer.Deserialize<List<Person>>(File.ReadAllText(ofd.FileName));

            using (var context = new GITZd2Entities())
            {
                foreach (var person in list)
                {
                    zad2_table zad = new zad2_table();
                    zad.ID = person.Id.ToString();
                    zad.OrderCode = person.CodeOrder.ToString();
                    zad.DateOfCreation = person.CreateDate.ToString();
                    zad.OrderTime = person.CreateTime.ToString();
                    zad.ClientCode = person.CodeClient.ToString();
                    zad.Services = person.Services.ToString();
                    zad.Status = person.Status.ToString();
                    zad.ClosingDate = person.ClosedDate.ToString();
                    zad.RentalTime = person.ProkatTime.ToString();


                    context.zad2_table.Add(zad);
                }
                context.SaveChanges();
                MessageBox.Show("Успешный импорт");
            }
        }

        class Person
        {
            public int Id { get; set; }
            public string CodeOrder { get; set; }
            public string CreateDate { get; set; }
            public string CreateTime { get; set; }
            public string CodeClient { get; set; }
            public string Services { get; set; }
            public string Status { get; set; }
            public string ClosedDate { get; set; }
            public string ProkatTime { get; set; }
        }

        private void Export_JSON_Click(object sender, RoutedEventArgs e)
        {

            using (var context = new GITZd2Entities())
            {
                ExportToWord(context.zad2_table.ToList());
            }
        }

        public static void ExportToWord(List<zad2_table> dataList)
        {
            // Group the objects by their "Status" property
            var groupedData = dataList.GroupBy(d => d.Status);

            // Create a new Word document
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            // Add a heading to the document
            foreach (var group in groupedData)
            {
                // Add a heading for the group
                Word.Paragraph heading = wordDoc.Content.Paragraphs.Add();
                heading.Range.Text = group.Key;
                heading.Range.Font.Bold = 1;
                heading.Range.Font.Size = 14;
                heading.Format.SpaceAfter = 24;
                heading.Range.InsertParagraphAfter();

                // Create a new table with headers
                Word.Range range = wordDoc.Content;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.Table table = wordDoc.Tables.Add(range, group.Count() + 1, 8);

                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "ID";
                table.Cell(1, 2).Range.Text = "OrderCode";
                table.Cell(1, 3).Range.Text = "DateOfCreation";
                table.Cell(1, 4).Range.Text = "OrderTime";
                table.Cell(1, 5).Range.Text = "ClientCode";
                table.Cell(1, 6).Range.Text = "Services";
                table.Cell(1, 7).Range.Text = "Status";
                table.Cell(1, 8).Range.Text = "ClosingDate";

                // Add data to the table
                int row = 2;
                foreach (var item in group)
                {
                    table.Cell(row, 1).Range.Text = item.ID;
                    table.Cell(row, 2).Range.Text = item.OrderCode;
                    table.Cell(row, 3).Range.Text = item.DateOfCreation;
                    table.Cell(row, 4).Range.Text = item.OrderTime;
                    table.Cell(row, 5).Range.Text = item.ClientCode;
                    table.Cell(row, 6).Range.Text = item.Services;
                    table.Cell(row, 7).Range.Text = item.Status;
                    table.Cell(row, 8).Range.Text = item.ClosingDate;
                    row++;
                }

                // Add a heading to the table
                table.Rows[1].HeadingFormat = -1;
                table.Rows[1].Range.Font.Bold = 1;
                table.Rows[1].Range.Font.Size = 12;
                table.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

                wordDoc.Content.InsertParagraphAfter();
            }

            // Save and close the document
            wordDoc.SaveAs2("C:\\Users\\azati\\OneDrive\\Рабочий стол\\DataGroupedByStatus.docx");
            wordDoc.Close();
            wordApp.Quit();

            MessageBox.Show("Успешный экспорт");
        }

    }
}
