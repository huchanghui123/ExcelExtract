using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelExtract
{
    public class Utils
    {
        public static void Test(List<Order> orderList)
        {
            Dictionary<string, List<Order>> result = new Dictionary<string, List<Order>>();
            
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

            Dictionary<string, List<Order>> result_new = new Dictionary<string, List<Order>>();
            foreach (string key in result.Keys)
            {
                List<Order> list = result[key];
                List<Order> list1 = new List<Order>();
                int i = 0;
                foreach(Order order in list)
                {
                    i += Convert.ToInt16(order.Num);
                    order.Num = i.ToString();
                    list1.Add(order);
                }

                list1.RemoveRange(0, list1.Count-1);

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

            //foreach(string key in result_new.Keys)
            //{
            //    Console.WriteLine("key:{0} count:{1}", key, result_new[key].Count);
            //    List<Order> list = result_new[key];
            //    foreach(Order order in list)
            //    {
            //        Console.WriteLine("key:{0} Model:{1} Num:{2}", key, order.Model, order.Num);
            //    }
            //}
            Save2Xlsx(result_new);
        }

        public static void  Save2Xlsx(Dictionary<string, List<Order>> order_dic)
        {
            //Workbook wb = new Workbook();
            //清除默认的工作表
            //wb.Worksheets.Clear();
            List<String> sheetlist = new List<String>();
            foreach (string key in order_dic.Keys)
            {
                List<Order> orderlist = order_dic[key];
                string sheet_name = orderlist.First().Client.Substring(0, 5);
                if(sheetlist.IndexOf(sheet_name) > 0)
                {
                    sheet_name += "1";
                }
                sheetlist.Add(sheet_name);
                //Worksheet st = wb.Worksheets.Add(sheet_name);
                //Console.WriteLine("sheet_name:{0}", sheet_name);
                foreach (Order order in orderlist)
                {
                    Console.WriteLine("key:{0} Model:{1} Num:{2}", key, order.Model, order.Num);

                }
            }

            //wb.SaveToFile("TEST销售出库单_new.xlsx", FileFormat.Version2013);
        }


        public static void FormatXlsx(List<Order> orderList)
        {
            Workbook wb = new Workbook();
            //清除默认的工作表
            wb.Worksheets.Clear();
            Worksheet st = wb.Worksheets.Add("TEST出库单");
            //创建样式
            CellStyle style = wb.Styles.Add("newStyle");
            style.Font.FontName = "宋体";
            //定义字体大小
            style.Font.Size = 12;
            st.ApplyStyle(style);

            //设置列宽
            st.Columns[0].ColumnWidth = 13F;
            st.Columns[1].ColumnWidth = 13F;
            st.Columns[2].ColumnWidth = 13F;
            st.Columns[3].ColumnWidth = 10F;
            st.Columns[4].ColumnWidth = 10F;
            st.Columns[5].ColumnWidth = 11F;
            st.Columns[6].ColumnWidth = 7F;

            //横向合并A1到G1的单元格
            st.Range["A1:G1"].Merge();
            st.Rows[0].RowHeight = 19F;
            //写入数据到A1单元格，设置文字格式及对齐方式
            //创建字体
            ExcelFont font1 = wb.CreateFont();
            font1.FontName = "宋体";
            font1.IsBold = true;
            font1.Size = 14;
            font1.Underline = FontUnderlineType.Single;

            //为A1单元格写入数据并设置字体
            RichText richText = st.Range["A1"].RichText;
            richText.Text = "出入库单";
            st.Range["A1"].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A1"].VerticalAlignment = VerticalAlignType.Center;
            richText.SetFont(0, richText.Text.ToArray().Length - 1, font1);

            st.Range["E2"].Value = "页码：";
            st.Range["F2"].Value = "第1页，共1页";

            
            st.Range["B3"].Value = "日期：";
            st.Range["C3"].Value = "2019-09-16";
            st.Range["E3"].Value = "单号：";
            st.Range["F3"].Value = "I0-2019-07-0012";

            st.Range["A4"].Value = "客户名称：";
            st.Range["D4"].Value = "单据类型：";
            st.Range["E4"].Value = "销售出库单";

            st.Range["F5:G5"].Merge();
            st.Range["F6:G6"].Merge();
            st.Range["F7:G7"].Merge();
            st.Range["F8:G8"].Merge();
            st.Range["F9:G9"].Merge();
            st.Range["F10:G10"].Merge();
            st.Range["F11:G11"].Merge();
            st.Range["F12:G12"].Merge();
            st.Range["F13:G13"].Merge();
            st.Range["F14:G14"].Merge();

            //创建一个DataTable
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("序号");
            dt1.Columns.Add("货品名称");
            dt1.Columns.Add("规格");
            dt1.Columns.Add("单位");
            dt1.Columns.Add("数量");
            dt1.Columns.Add("备注");
            dt1.Rows.Add("1", "工控主机", "Q190S", "台", "10", "");
            dt1.Rows.Add("2", "工控主机", "Q220S", "台", "2", "");
            dt1.Rows.Add("");
            dt1.Rows.Add("");
            dt1.Rows.Add("");
            dt1.Rows.Add("");
            dt1.Rows.Add("");
            dt1.Rows.Add("");
            st.Range["A"+14+":"+"D"+14].Merge();
            st.Range["A" + 14 + ":" + "D" + 14].HorizontalAlignment = HorizontalAlignType.Right;
            dt1.Rows.Add("合计：","","","", "12");

            st.Range["A5:G13"].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A14"].HorizontalAlignment = HorizontalAlignType.Right;
            st.Range["E14"].HorizontalAlignment = HorizontalAlignType.Center;

            //设置网格线样式及颜色
            st.Range["A5:G14"].BorderAround(LineStyleType.Thin);
            st.Range["A5:G14"].BorderInside(LineStyleType.Thin);
            st.Range["A5:G14"].Borders.KnownColor = ExcelColors.Black;

            //将DataTable数据写入工作表
            st.InsertDataTable(dt1, true, 5, 1, true);
            int index = 5 + dt1.Rows.Count + 2;

            st.Range["A"+ index].Value = "部门：";
            st.Range["B"+ index].Value = "业务员：";
            st.Range["D" + index].Value = "制单人：";
            st.Range["F" + index].Value = "审核人：";
            
            wb.SaveToFile("TEST销售出库单.xlsx", FileFormat.Version2013);
            
        }
    }
}
