using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Data;
using System.IO;


namespace excel_to_db
{
    class ExcelHelper
    {
        public IEnumerable<DataRow> Read(string filePath)
        {
            IEnumerable<DataRow> dataRows = null;
            using (FileStream fs = File.Open(filePath,
                FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = null;
                DataTable worksheet = null;
                if (filePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(fs);
                }
                else if (filePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                }
                else
                {
                    return null;
                }
                //reader.IsFirstRowAsColumnNames = true;
                DataSet dataSet = reader.AsDataSet();
                /*
                int count = 0;
                foreach (var table in dataSet.Tables)
                {
                    count++;
                }
                Console.WriteLine("---------" + count);
                */
                //int countTables = dataSet.Tables.Count();
                worksheet = dataSet.Tables[0];
                dataRows = from DataRow dr in worksheet.Rows
                           select dr;
                reader.Close();
            }
            return dataRows;
        }
    }
}
