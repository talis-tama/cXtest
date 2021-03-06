﻿using System;
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
            /*using(var book=new XLWorkbook(@"C:\Users\"+Environment.UserName+"\\Desktop\\test.xlsx",XLEventTracking.Disabled))
            {
                var sheet1 = book.Worksheet(1);
                Console.WriteLine(sheet1.Name);//debug
                var cell = sheet1.Cell("A1");
                var cell1 = sheet1.Cell("B2");
                string dat;
                dat = cell.GetString();
                cell1.Value = dat;
                book.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
            }//A1からB2へコピー*/
            using(var book=new XLWorkbook(@"C:\Users\" + Environment.UserName + "\\Desktop\\test.xlsx", XLEventTracking.Disabled))
            {
                var sheet1 = book.Worksheet(1);
                Console.WriteLine(sheet1.Name);//debug
                string[] dat = null;
                int a, b, d, e;
                int c = 0;
                int f = 0;
                for(a = 1; a <= 3; a++)//A→C
                {
                    for(b = 1; b <= 4; b++)//1→4
                    {
                        Array.Resize(ref dat, c + 1);//この書き方はまずい気がする
                        var cell = sheet1.Cell(b, a);//数字にて対象セル指定
                        dat[c] = cell.GetString();
                        Console.WriteLine(dat[c]);//debug
                        c++;
                    }
                }
                for (d = 6; d <= 8; d++)//F→H
                {
                    for (e = 6; e <= 9; e++)//6→9
                    {
                        var cell1 = sheet1.Cell(e, d);
                        cell1.Value = dat[f];
                        f++;
                    }
                }
                book.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
            }//A1→C4からF6→H9へコピー(範囲指定読み出し、書き込みテスト)
            /*for(int a = 1; a <= 5; a++)
            {
                for(int b = 1; b <= 7; b++)
                {
                    Console.WriteLine(a.ToString() + b.ToString());
                }
            }//仕様忘れ確認*/
        }
    }
}
