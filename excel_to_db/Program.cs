using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;


namespace excel_to_db
{
    class Program
    {
        static void Main(string[] args)
        {
            string conStr = ConfigurationManager.AppSettings["ConnectionString"];

            ReadExcel(conStr);
        }

        static void ReadExcel(string conToDb)
        {
            using (SqlConnection con = new SqlConnection(conToDb))
            {
                con.Open();
                Console.WriteLine("Conections succed Data Source=somebase.mssql.somee.com;Initial Catalog=somebase;User ID=finiuk_SQLLogin_1;");
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.Connection = con;

                ExcelHelper excelHelper = new ExcelHelper();
                var items = excelHelper.Read("data.xlsx");
                string[] users = { "", "", "" };
                int counter = 0;
                foreach (var item in items)
                {
                    var cols = item.ItemArray;
                    foreach (var itemCol in cols)
                    {
                        users[counter++] = itemCol.ToString();
                        //Console.Write($"{itemCol.ToString(),30}({++counter})");
                        if (counter % 3 == 0)
                        {
                            string query = "INSERT INTO [dbo].[tblUsers] ([Firstname],[Lastname],[Email]) " +
                            $"VALUES ('{users[0]}','{users[1]}','{users[2]}')";
                            sqlCommand.CommandText = query;
                            int countInsert = sqlCommand.ExecuteNonQuery();
                            Console.WriteLine($"{users[0], 15} {users[1], 15} {users[2], 30} ADDED");
                            counter = 0;
                        }
                    }
                    
                }
            }
        }
    }
}
