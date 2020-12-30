using Spire.Xls;
using System;
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
            StringBuilder sb = new StringBuilder();
            int num = 0;
            Console.WriteLine("Rows.Count:{0} Columns.Coun:{1}", dt.Rows.Count, dt.Columns.Count);
            //Console.ReadKey();

            Workbook wb = new Workbook();
            //清除默认的工作表
            wb.Worksheets.Clear();

            foreach (DataRow d in dt.Rows)
            {
                for(int i=0;i<dt.Columns.Count;i++)
                {
                    sb.Append(d[i].ToString()+" ");
                }
                //if(d[1].ToString().Length>0)
                //{
                //Worksheet sheet1 = wb.Worksheets.Add("hahahah"+num);
                //sheet1.Range["A1"].Value = sb.ToString();
                //}

                sb.Append("\r\n");
                Console.WriteLine("no:{0} ====>>> sb:{1}", num, sb.ToString());
                num++;
                sb.Clear();
            }

            Worksheet st = wb.Worksheets.Add("TEST出库单");
            //横向合并A1到G1的单元格
            st.Range["A1:G1"].Merge();
            st.Rows[0].RowHeight = 22F;
            //写入数据到A1单元格，设置文字格式及对齐方式
            st.Range["A1"].Value = "深圳市千度科技有限公司";
            st.Range["A1"].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A1"].VerticalAlignment = VerticalAlignType.Center;
            st.Range["A1"].Style.Font.IsBold = true;
            st.Range["A1"].Style.Font.Size = 16F;

            st.Range["A2:G2"].Merge();
            st.Rows[1].RowHeight = 20F;
            st.Range["A2"].Value = "销售出库单";
            st.Range["A2"].HorizontalAlignment = HorizontalAlignType.Center;
            st.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
            st.Range["A2"].Style.Font.Size = 14F;

            st.Range["A3:C3"].Merge();
            st.Range["D3:E3"].Merge();
            st.Range["F3:G3"].Merge();
            st.Range["A3"].Value = "购货单位：苏州恩迪科技有限公司";
            st.Range["D3"].Value = "日期：2019-7-05";
            st.Range["F3"].Value = "编号：I0-2019-07-0012";

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
            st.Range["A"+12+":"+"E"+12].Merge();
            st.Range["A" + 12 + ":" + "E" + 12].HorizontalAlignment = HorizontalAlignType.Right;
            dt1.Rows.Add("合计：","","","","", "12");

            //设置网格线样式及颜色
            st.Range["A5:G12"].BorderAround(LineStyleType.Thin);
            st.Range["A5:G12"].BorderInside(LineStyleType.Thin);
            st.Range["A5:G12"].Borders.KnownColor = ExcelColors.Black;

            //将DataTable数据写入工作表
            st.InsertDataTable(dt1, true, 5, 1, true);
            Console.WriteLine("rows--------"+ dt1.Rows.Count);
            int index = 5 + dt1.Rows.Count + 2;

            st.Range["A"+ index].Value = "审核：";
            st.Range["B"+ index].Value = "陈蓉";
            st.Range["C" + index].Value = "发货：";
            st.Range["D" + index].Value = "陆海";
            st.Range["E" + index].Value = "制单：";
            st.Range["F" + index].Value = "赵静";

            wb.SaveToFile("TEST销售出库单.xlsx", FileFormat.Version2013);
            //wb.SaveToFile("创建Excel.xlsx", FileFormat.Version2013);

            Console.ReadKey();
        }
    }
}
