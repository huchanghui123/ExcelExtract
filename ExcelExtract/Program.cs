using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelExtract
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("输入文件路径：");
            //String file_path = Console.ReadLine().Replace("\"","");
            String file_path = "C:\\Users\\16838\\Desktop\\销项明细台账20201230 - 副本.xlsx";
            //Console.WriteLine("path:{0}", file_path);
            
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@file_path);
            //获取第一张sheet
            Worksheet sheet = workbook.Worksheets[0];
            //设置range范围
            CellRange range = sheet.Range[sheet.FirstRow+2, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn-1];
            //输出数据, 同时输出列名以及公式值
            DataTable dt = sheet.ExportDataTable(range, true, true);
            //Console.WriteLine("Rows.Count:{0} Columns.Coun:{1}", dt.Rows.Count, dt.Columns.Count);
            //Console.ReadKey();

            List<Order> orderList = new List<Order>();
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if(dr[4].ToString().Length > 0)
                {
                    Order order = new Order
                    {
                        Client = dr[4].ToString(),
                        Date = dr[3].ToString(),
                        No = i.ToString(),
                        Name = dr[6].ToString().Split('*')[2],
                        Model = dr[7].ToString(),
                        Unit = dr[8].ToString(),
                        Num = dr[9].ToString()
                    };
                    orderList.Add(order);
                    //Console.WriteLine(order.ToString());
                    i++;
                }
                
            }
            Utils.Test(orderList);
            //Utils.FormatXlsx(orderList);

            Console.ReadKey();
        }
    }
}
