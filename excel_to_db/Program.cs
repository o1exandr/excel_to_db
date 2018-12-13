using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;

namespace excel_to_db
{
    class Program
    {
        static void Main(string[] args)
        {
          
        }

        static void ReadExcel()
        {

            ExcelHelper excelHelper = new ExcelHelper();
            var items = excelHelper.Read("book.xlsx");
            foreach (var item in items)
            {
                var cols = item.ItemArray;
                foreach (var itemCol in cols)
                    Console.Write($"{itemCol.ToString()}  - ");
                Console.WriteLine();
            }
        }
        static void WriteExcel()
        {
            using (DataTable dt = new DataTable())
            {
                dt.Columns.Add(new DataColumn("Id",
                    Type.GetType("System.Int32")));
                dt.Columns.Add(new DataColumn("Name",
                    Type.GetType("System.String")));
                dt.Columns.Add(new DataColumn("Hobi",
                    Type.GetType("System.String")));
                dt.Columns.Add(new DataColumn("Image",
                    Type.GetType("System.String")));
                string fileName = "outputBooks.xlsx";
                using (XLWorkbook wb = new XLWorkbook())
                {
                    for (int i = 0; i < 10; i++)
                    {
                        object[] listCols =
                        {
                            i+1,
                            "Step",
                            null,
                            Guid.NewGuid().ToString()
                        };
                        dt.Rows.Add(listCols);
                    }
                    wb.Worksheets.Add(dt, "users-images");
                    wb.SaveAs(fileName);

                }

            }
        }
        static void Main(string[] args)
        {
            ReadExcel();
            //List<Person> list = new List<Person>
            //{
            //    new Person
            //    {
            //        FullName = "Іван Васильович",
            //        DateBirth=DateTime.Now
            //    }
            //};
            //WordHelper wordHelper = new WordHelper();
            //wordHelper.SavePersonToWord(list,"doc.docx");
        }
    }
}
