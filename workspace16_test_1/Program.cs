using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace workspace16_test_1
{
    class Program
    {
        static void Main(string[] args)
        {
            using(var book=new XLWorkbook(@"C:\Users\"+Environment.UserName+"\\Desktop\\test.xlsx",XLEventTracking.Disabled))
            {
                var sheet1 = book.Worksheet(1);
                Console.WriteLine(sheet1.Name);//debug
                var cell = sheet1.Cell("A1");
                var cell1 = sheet1.Cell("B2");
                string dat;
                dat = cell.GetString();
                cell1.Value = dat;
                book.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
            }
        }
    }
}
