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
           

        }
    }
}
