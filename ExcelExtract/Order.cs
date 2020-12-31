using System;

namespace ExcelExtract
{
    public class Order
    {
        private string client;
        private string date;
        private string no;
        private string name;
        private string model;
        private string unit;
        private string num;

        public string Client { get => client; set => client = value; }
        public string Date { get => date; set => date = value; }
        public string No { get => no; set => no = value; }
        public string Name { get => name; set => name = value; }
        public string Model { get => model; set => model = value; }
        public string Unit { get => unit; set => unit = value; }
        public string Num { get => num; set => num = value; }

        public Order() { }

        public Order(string client, string date, string no, string name, string model, string unit, string num)
        {
            this.Client = client;
            this.Date = date;
            this.No = no;
            this.Name = name;
            this.Model = model;
            this.Unit = unit;
            this.Num = num;
        }

        public override string ToString()
        {
            string str = String.Format("no:{0} \r\n" + 
                "date:{1} \r\n" +
                "client:{2} \r\n" +
                "name:{3} \r\n" +
                "model:{4} \r\n" +
                "unit:{5} \r\n" +
                "num:{6} ", No, Date, Client, Name, Model, Unit, Num);
            return str;
        }
    }
}
