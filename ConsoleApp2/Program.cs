using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application excelApp = new Application();
            //Проверка на наличие excel
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            //открывает файл по пути
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\QWERTY\Desktop\0.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            //Считывает строки и колонки
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                //Вывод в консоли в заданном порядке (с помощью разделителей строк)
                Console.Write("\r\n");
                for (int j = 1; j <= cols; j++)
                {


                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t\n");
                }
            }
            //Выход из excel и вывод в консоль
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();
        }
    }
}
