using OpenXml_Excel;
using System.Data;
using System.Drawing;
using System.Reflection;

namespace UnitTest
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            string text = Properties.Resources.table1;

            DataTable dt = new DataTable();

            dt.Columns.Add(new DataColumn("Name", typeof(string)));
            dt.Columns.Add(new DataColumn("Year", typeof(string)));
            dt.Columns.Add(new DataColumn("Month", typeof(string)));
            dt.Columns.Add(new DataColumn("Desc", typeof(string)));

            int colLen = 4;
            string[] lines = text.Split("\r\n");
            foreach (string line in lines)
            {
                string[] array = line.Split("\t");
                DataRow row = dt.NewRow();
                for (int i = 0; i < colLen; i++)
                {
                    row[i] = array[i];
                }
                dt.Rows.Add(row);
            }

            ExcelHelper.Write(@"d:/log/table1.xlsx", dt);
            Assert.Pass();
        }

        [Test]
        public void Test2()
        {
            string text = Properties.Resources.table1;

            DataTable dt = new DataTable();

            dt.Columns.Add(new DataColumn("Name", typeof(string)));
            dt.Columns.Add(new DataColumn("Year", typeof(string)));
            dt.Columns.Add(new DataColumn("Month", typeof(string)));
            dt.Columns.Add(new DataColumn("Desc", typeof(string)));
            dt.Columns.Add(new DataColumn("Image", typeof(string)));

            int colLen = 4;
            string[] lines = text.Split("\r\n");
            foreach (string line in lines)
            {
                string[] array = line.Split("\t");
                DataRow row = dt.NewRow();
                for (int i = 0; i < colLen; i++)
                {
                    row[i] = array[i];
                }
                row[4] = @"D:\log\dog\dog1.png";
                dt.Rows.Add(row);
            }

            try
            {
                ExcelHelper.Write(@"d:/log/table1.xlsx", dt);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            Assert.Pass();
        }


    }
}