using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excsv
{
    public static class ExcelFileExtension
    {
        public static bool SaveAsCSV(this string excel_path, string csv_path)
        {
            using (var stream = new FileStream(excel_path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = null;
                if (excel_path.EndsWith(".xls"))
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                else if (excel_path.EndsWith("xlsx") || excel_path.EndsWith("xlsm"))
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                if (reader == null) return false;
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tr) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });
                var csvContent = new StringBuilder();
                for (var row_no = 0; row_no < ds.Tables[0].Rows.Count; row_no++)
                {
                    var lineContent = new List<string>();
                    for (var col_no = 0; col_no < ds.Tables[0].Columns.Count; col_no++)
                    {
                        lineContent.Add(ds.Tables[0].Rows[row_no][col_no].ToString());
                    }
                    csvContent.AppendLine(string.Join(",", lineContent));
                }
                File.WriteAllText(csv_path, csvContent.ToString(), Encoding.UTF8);
                //var csv = new StreamWriter(csv_path, false);
                //csv.Write(csvContent);
                //csv.Close();
                return true;
            }
        }
    }
}
