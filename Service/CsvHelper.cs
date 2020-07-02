using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Service
{
    public static class CsvHelper
    {
        public static void DataTable2Csv(DataTable table, string filePath = "", bool isHeader = true)
        {
            var lines = new List<string>();
            filePath = string.IsNullOrEmpty(filePath) ? DateTime.Now.ToString("yyyyMMdd_HHmmss") : filePath;

            try
            {
                if (Path.GetExtension(filePath) != ".csv")
                {
                    filePath += ".csv";
                }

                string[] columnNames = table.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToArray();

                string header = string.Join(",", columnNames);
                lines.Add(header);

                var values = table.AsEnumerable()
                    .Select(row => string.Join(",", row.ItemArray));

                lines.AddRange(values);

                File.WriteAllLines(filePath, lines);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable Csv2DataTable(string filePath, char separator = ',', bool isHeader = true)
        {
            DataTable table = new DataTable();

            try
            {
                StreamReader streamReader = new StreamReader(filePath);

                if (isHeader)
                {
                    string[] headers = streamReader.ReadLine().Split(separator);
                    foreach (string header in headers)
                    {
                        table.Columns.Add(header);
                    }
                }

                while (!streamReader.EndOfStream)
                {
                    table.Rows.Add(streamReader.ReadLine().Split(','));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return table;
        }

        public static DataTable Csv2DataTable(string filePath, char[] separators, bool isHeader = true)
        {
            DataTable table = new DataTable();

            try
            {
                StreamReader streamReader = new StreamReader(filePath);

                if (isHeader)
                {
                    string[] headers = streamReader.ReadLine().Split(separators);
                    foreach (string header in headers)
                    {
                        table.Columns.Add(header);
                    }
                }

                while (!streamReader.EndOfStream)
                {
                    table.Rows.Add(streamReader.ReadLine().Split(','));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
            return table;
        }
    }
}
