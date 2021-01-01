using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExcelExtract
{
    public class Utils
    {
        public static void Test(string file_name, List<Order> orderList)
        {
            Dictionary<string, List<Order>> result = new Dictionary<string, List<Order>>();
            //通过日期、客户名、型号筛选订单
            //同日期、客户、型号判断为一个出库单
            foreach(Order order in orderList)
            {
                string key = order.Date + order.Client +"-"+ order.Model;
                
                if (result.ContainsKey(key))
                {
                    List<Order> order1 = result[key];
                    order1.Add(order);
                }
                else
                {
                    List<Order> order1 = new List<Order>();
                    order1.Add(order);
                    result.Add(key, order1);
                }
            }
            //将出库单内相同型号的订单合并，数量相加
            //重新根据日期、客户名筛选订单
            //同日期、客户判断为一个出库单
            Dictionary<string, List<Order>> result_new = new Dictionary<string, List<Order>>();
            foreach (string key in result.Keys)
            {
                List<Order> list = result[key];
                List<Order> list1 = new List<Order>();
                int i = 0;
                foreach(Order order in list)
                {
                    i += Convert.ToInt16(order.Num);
                    if(i<=0)
                    {
                        list1.Add(order);
                    }
                    else
                    {
                        order.Num = i.ToString();
                        list1.Add(order);
                    }
                    //Console.WriteLine("key:{0} Model:{1} Num:{2}", key, order.Model, order.Num);
                }
                //取消订单数量会<=0，将数据保留 
                if(i>0)
                {
                    list1.RemoveRange(0, list1.Count-1);
                }

                //key:日期+客户名
                string key_new = key.Split('-')[0];
                if(result_new.ContainsKey(key_new))
                {
                    List<Order> order_list = result_new[key_new];
                    order_list.AddRange(list1);
                    result_new[key_new] = order_list;
                }
                else
                {
                    result_new.Add(key_new, list1);
                }
            }

            //foreach (string key in result_new.Keys)
            //{
            //    Console.WriteLine("key:{0} count:{1}", key, result_new[key].Count);
            //    List<Order> list = result_new[key];
            //    foreach (Order order in list)
            //    {
            //        Console.WriteLine("key:{0} Model:{1} Num:{2}", key, order.Model, order.Num);
            //    }
            //}
            Save2Xlsx(file_name, result_new);
        }

        public static void  Save2Xlsx(string fileName, Dictionary<string, List<Order>> order_dic)
        {
            Workbook wb = new Workbook();
            //清除默认的工作表
            wb.Worksheets.Clear();
            List<String> sheetlist = new List<String>();
            int no = 1;
            foreach (string key in order_dic.Keys)
            {
                List<Order> orderlist = order_dic[key];
                string sn = orderlist.First().Client;
                string sheet_name = sn;
                if (sn.Length >= 5)
                {
                    sheet_name = sn.Substring(0, 5);
                }
                
                if(sheetlist.IndexOf(sn) >0)
                {
                    int _index = sheetlist.FindAll((String str) => str == sn).Count;
                    //重复客户，页签名+1
                    sheet_name += _index.ToString();
                }
                sheetlist.Add(sn);

                Worksheet st = wb.Worksheets.Add(sheet_name);
                //创建样式
                CellStyle style = wb.Styles.Add("newStyle");
                style.Font.FontName = "宋体";
                style.Font.Size = 12;
                st.ApplyStyle(style);
                //创建字体
                ExcelFont font1 = wb.CreateFont();
                font1.FontName = "宋体";
                font1.IsBold = true;
                font1.Size = 14;
                font1.Underline = FontUnderlineType.Single;

                Console.WriteLine("sheet_name:{0}", sheet_name);
                foreach (Order order in orderlist)
                {
                    Console.WriteLine("key:{0} Model:{1} Num:{2}", key, order.Model, order.Num);
                }
                FormatXlsx(no, st, font1, orderlist);
                no++;
            }

            wb.SaveToFile(fileName+"_销售出库单.xlsx", FileFormat.Version2013);
        }


        public static void FormatXlsx(int no, Worksheet st, ExcelFont font, List<Order> orderList)
        {
            string no_str = no.ToString().PadLeft(4, '0');
            Order order = orderList.First();
            string date = order.Date.Replace('年', '-').Replace('月', '-').Replace('日',' ').Trim();
            //设置列宽
            st.Columns[0].ColumnWidth = 11F;
            st.Columns[1].ColumnWidth = 14F;
            st.Columns[2].ColumnWidth = 15F;
            st.Columns[3].ColumnWidth = 10F;
            st.Columns[4].ColumnWidth = 12F;
            st.Columns[5].ColumnWidth = 9F;
            st.Columns[6].ColumnWidth = 11F;

            //横向合并A1到G1的单元格
            st.Range["A1:G1"].Merge();
            st.Rows[0].RowHeight = 22F;
            //写入数据到A1单元格，设置文字格式及对齐方式
            

            //为A1单元格写入数据并设置字体
            RichText richText = st.Range["A1"].RichText;
            richText.Text = "出入库单";
            st.Range["A1"].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A1"].VerticalAlignment = VerticalAlignType.Center;
            richText.SetFont(0, richText.Text.ToArray().Length - 1, font);

            st.Range["E2"].Value = "页码：";
            st.Range["E2"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["F2"].Value = "第1页，共1页";
            st.Range["F2:G2"].Merge();


            st.Range["B3"].Value = "日期：";
            st.Range["B3"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["C3"].Value = date;
            st.Range["E3"].Value = "单号：";
            st.Range["E3"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["F3"].Value = "I0-" + date + "-" + no_str;
            st.Range["F3:G3"].Merge();

            st.Range["A4"].Value = "客户名称：";
            st.Range["A4"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["B4"].Value = order.Client;
            st.Range["B4"].HorizontalAlignment = HorizontalAlignType.Left;
            st.Range["E4"].Value = "单据类型：";
            st.Range["E4"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["F4"].Value = "销售出库单";
            st.Range["B4:D4"].Merge();
            st.Range["F4:G4"].Merge();

            //创建一个DataTable
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("序号");
            dt1.Columns.Add("货品名称");
            dt1.Columns.Add("规格");
            dt1.Columns.Add("单位");
            dt1.Columns.Add("数量");
            dt1.Columns.Add("备注");
            int length = orderList.Count;
            int i = 1;
            int total = 0;
            foreach(Order o in orderList)
            {
                dt1.Rows.Add(i.ToString(), o.Name, o.Model, o.Unit, o.Num, "");
                total += Convert.ToInt16(o.Num);
                i++;
            }
            int j = 8 - length;
            if(j > 0)
            {
                while(j-->0)
                {
                    dt1.Rows.Add("");
                }
            }
            for(int k=5;k< dt1.Rows.Count+7;k++)
            {
                st.Range["F"+k+":G"+k].Merge();
            }
            //Console.WriteLine("dt1.Rows.Count:{0}", dt1.Rows.Count);
            int index = dt1.Rows.Count + 6;
            st.Range["A"+ index + ":"+"D"+ index].Merge();
            st.Range["A" + index + ":" + "D" + index].HorizontalAlignment = HorizontalAlignType.Right;
            dt1.Rows.Add("合计：","","","", total.ToString());

            st.Range["A5:G"+(index-1)].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A"+ index].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["E"+ index].HorizontalAlignment = HorizontalAlignType.Center;

            //设置网格线样式及颜色
            st.Range["A5:G"+ index].BorderAround(LineStyleType.Thin);
            st.Range["A5:G"+ index].BorderInside(LineStyleType.Thin);
            st.Range["A5:G"+ index].Borders.KnownColor = ExcelColors.Black;

            //将DataTable数据写入工作表
            st.InsertDataTable(dt1, true, 5, 1, true);

            st.Range["A"+ (index+2)].Value = "审核：";
            st.Range["A" + (index + 2)].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["B"+ (index + 2)].Value = "陈蓉";

            st.Range["C"+ (index + 2)].Value = "发货：";
            st.Range["C" + (index + 2)].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["D" + (index + 2)].Value = "陆海";

            st.Range["F" + (index + 2)].Value = "制单：";
            st.Range["F" + (index + 2)].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["G" + (index + 2)].Value = "赵静";
            
        }
    }
}
