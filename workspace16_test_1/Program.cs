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
            string dat;
            var vb = new XLWorkbook(@"C:\test.xlsx");
            var cell = vb.Cell("A1");//A1を読み込み
            dat = cell.Value.ToString();
            vb.Cell("B2").Value = dat;//B2へ書き込み
            vb.SaveAs(@"C:\testafter.xlsx");
        }
    }
}
