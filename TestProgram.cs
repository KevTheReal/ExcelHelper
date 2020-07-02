using ExcelHelper.Service;
using System;
using System.Data;

namespace ExcelHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DataTable table = NpoiExcelHelper.Excel2DataTable("Sample.xls");
                CsvHelper.DataTable2Csv(table);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
