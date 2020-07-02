using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace ExcelHelper.Service
{
    /// <summary>
    /// Converts between Excels and DataSets using Nopi.
    /// </summary>
    public static class NpoiExcelHelper
    {
        /// <summary>
        /// Converts DataSet to Excel.
        /// </summary>
        /// <param name="ds">The DataSet to be converted.</param>
        /// <param name="filePath">The path and name of the file to be exported.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        /// <returns></returns>
        public static void DataSet2Excel(DataSet ds, string filePath = "", bool isHeader = true)
        {
            int sheetIndex = 0;
            int startRow = isHeader ? 1 : 0;
            IWorkbook workbook = null;

            filePath = string.IsNullOrEmpty(filePath) ? DateTime.Now.ToString("yyyyMMdd_HHmmss") : filePath;

            try
            {
                string extension = Path.GetExtension(filePath);

                if (extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook();
                }
                else if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    filePath = filePath + ".xlsx";
                    workbook = new XSSFWorkbook();
                }

                foreach (DataTable table in ds.Tables)
                {
                    sheetIndex++;

                    if (table != null && table.Rows.Count > 0)
                    {
                        ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(table.TableName) ? ("sheet" + sheetIndex.ToString()) : table.TableName);
                        IRow row = null;
                        ICell cell = null;

                        if (isHeader)
                        {
                            row = sheet.CreateRow(0);
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                cell = row.CreateCell(j);
                                cell.SetCellValue(table.Columns[j].ColumnName);
                            }
                        }

                        for (int i = startRow; i < table.Rows.Count; i++)
                        {
                            row = sheet.CreateRow(i);
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                cell = row.CreateCell(j);
                                cell.SetCellValue(table.Rows[i][j].ToString());
                            }
                        }
                    }
                }

                using (FileStream stream = File.Create(filePath))
                {
                    workbook.Write(stream);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts DataTable to Excel
        /// </summary>
        /// <param name="table">The DataTable to be converted.</param>
        /// <param name="filePath">The path and name of the file to be exported.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        public static void DataTable2Excel(DataTable table, string filePath = "", bool isHeader = true)
        {
            int startRow = isHeader ? 1 : 0;
            IWorkbook workbook = null;

            filePath = string.IsNullOrEmpty(filePath) ? DateTime.Now.ToString("yyyyMMdd_HHmmss") : filePath;

            try
            {
                string extension = Path.GetExtension(filePath);

                if (extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook();
                }
                else if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    filePath = filePath + ".xlsx";
                    workbook = new XSSFWorkbook();
                }

                if (table != null && table.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(table.TableName) ? "sheet" : table.TableName);
                    IRow row = null;
                    ICell cell = null;

                    if (isHeader)
                    {
                        row = sheet.CreateRow(0);
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            cell = row.CreateCell(j);
                            cell.SetCellValue(table.Columns[j].ColumnName);
                        }
                    }

                    for (int i = startRow; i < table.Rows.Count; i++)
                    {
                        row = sheet.CreateRow(i);
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            cell = row.CreateCell(j);
                            cell.SetCellValue(table.Rows[i][j].ToString());
                        }
                    }
                }


                using (FileStream stream = File.Create(filePath))
                {
                    workbook.Write(stream);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts Excel to DataSet.
        /// </summary>
        /// <param name="filePath">The path and name of the file to be converted.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        /// <returns></returns>
        public static DataSet Excel2DataSet(string filePath, bool isHeader = true)
        {
            DataSet dataSet = new DataSet();
            IWorkbook workbook = null;
            int startRow = isHeader ? 1 : 0;

            try
            {
                string extension = Path.GetExtension(filePath);
                FileStream fileStream = File.OpenRead(filePath);

                // 2007
                if (extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                // 2003
                else if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else
                {
                    throw new ArgumentException("The type of imported file is invalid. This function only accepts Excels(xls/xlsx).", "fileName");
                }

                if (workbook != null)
                {
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        ISheet sheet = workbook.GetSheetAt(i);
                        if (sheet == null)
                        {
                            continue;
                        }

                        IRow row = null;
                        ICell cell = null;

                        DataTable table = new DataTable(sheet.SheetName);

                        if (isHeader)
                        {
                            row = sheet.GetRow(0);
                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                cell = row.GetCell(k);
                                table.Columns.Add(cell.StringCellValue);
                            }
                        }

                        for (int j = startRow; j <= sheet.LastRowNum; ++j)
                        {
                            row = sheet.GetRow(j);
                            if (row == null)
                            {
                                continue;
                            }

                            DataRow dataRow = table.NewRow();

                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                cell = row.GetCell(k);
                                if (cell == null)
                                {
                                    dataRow[k] = string.Empty;
                                }
                                else
                                {
                                    // CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                    switch (cell.CellType)
                                    {
                                        case CellType.Numeric:
                                            dataRow[k] = cell.NumericCellValue;
                                            break;
                                        case CellType.String:
                                            dataRow[k] = cell.StringCellValue;
                                            break;
                                        case CellType.Formula:
                                            dataRow[k] = cell.StringCellValue;
                                            break;
                                        case CellType.Boolean:
                                            dataRow[k] = cell.BooleanCellValue;
                                            break;
                                        default:
                                            dataRow[k] = string.Empty;
                                            break;
                                    }
                                }
                            }

                            table.Rows.Add(dataRow);
                        }

                        dataSet.Tables.Add(table);
                    }
                }

                return dataSet;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts Excel to DataSet.
        /// </summary>
        /// <param name="stream">The memory stream to be converted.</param>
        /// <param name="contentType">The mime content type.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        /// <returns></returns>
        public static DataSet Excel2DataSet(MemoryStream stream, string contentType, bool isHeader = true)
        {
            DataSet dataSet = new DataSet();
            IWorkbook workbook = null;
            int startRow = isHeader ? 1 : 0;

            try
            {
                // 2007
                if (contentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    workbook = new XSSFWorkbook(stream);
                }
                // 2003
                else if (contentType == "application/vnd.ms-excel")
                {
                    workbook = new HSSFWorkbook(stream);
                }
                else
                {
                    throw new ArgumentException("The type of imported file is invalid. This function only accepts Excels(xls/xlsx).", "fileName");
                }

                if (workbook != null)
                {
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        ISheet sheet = workbook.GetSheetAt(i);
                        if (sheet == null)
                        {
                            continue;
                        }

                        IRow row = null;
                        ICell cell = null;

                        DataTable table = new DataTable(sheet.SheetName);

                        if (isHeader)
                        {
                            row = sheet.GetRow(0);
                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                cell = row.GetCell(k);
                                table.Columns.Add(cell.StringCellValue);
                            }
                        }

                        for (int j = startRow; j <= sheet.LastRowNum; ++j)
                        {
                            row = sheet.GetRow(j);
                            if (row == null)
                            {
                                continue;
                            }

                            DataRow dataRow = table.NewRow();

                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                cell = row.GetCell(k);
                                if (cell == null)
                                {
                                    dataRow[k] = string.Empty;
                                }
                                else
                                {
                                    // CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                    switch (cell.CellType)
                                    {
                                        case CellType.Numeric:
                                            dataRow[k] = cell.NumericCellValue;
                                            break;
                                        case CellType.String:
                                            dataRow[k] = cell.StringCellValue;
                                            break;
                                        case CellType.Formula:
                                            dataRow[k] = cell.StringCellValue;
                                            break;
                                        case CellType.Boolean:
                                            dataRow[k] = cell.BooleanCellValue;
                                            break;
                                        default:
                                            dataRow[k] = string.Empty;
                                            break;
                                    }
                                }
                            }

                            table.Rows.Add(dataRow);
                        }

                        dataSet.Tables.Add(table);
                    }
                }

                return dataSet;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts Excel to DataTable.
        /// </summary>
        /// <param name="filePath">The path and name of the file to be converted.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        /// <returns></returns>
        public static DataTable Excel2DataTable(string filePath, bool isHeader = true, int sheetNo = 0)
        {
            DataTable table = null;
            IWorkbook workbook = null;
            int startRow = isHeader ? 1 : 0;

            try
            {
                string extension = Path.GetExtension(filePath);
                FileStream fileStream = File.OpenRead(filePath);

                // 2007
                if (extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                // 2003
                else if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else
                {
                    throw new ArgumentException("The type of imported file is invalid. This function only accepts Excels(xls/xlsx).", "fileName");
                }

                if (workbook != null)
                {
                    ISheet sheet = workbook.GetSheetAt(sheetNo);

                    if (sheet == null)
                    {
                        throw new NullReferenceException("This sheet is null.");
                    }

                    IRow row = null;
                    ICell cell = null;

                    table = new DataTable(sheet.SheetName);

                    if (isHeader)
                    {
                        row = sheet.GetRow(0);
                        for (int k = 0; k < row.LastCellNum; k++)
                        {
                            cell = row.GetCell(k);
                            table.Columns.Add(cell.StringCellValue);
                        }
                    }

                    for (int j = startRow; j <= sheet.LastRowNum; ++j)
                    {
                        row = sheet.GetRow(j);
                        if (row == null)
                        {
                            continue;
                        }

                        DataRow dataRow = table.NewRow();

                        for (int k = 0; k < row.LastCellNum; k++)
                        {
                            cell = row.GetCell(k);
                            if (cell == null)
                            {
                                dataRow[k] = string.Empty;
                            }
                            else
                            {
                                // CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        dataRow[k] = cell.NumericCellValue;
                                        break;
                                    case CellType.String:
                                        dataRow[k] = cell.StringCellValue;
                                        break;
                                    case CellType.Formula:
                                        dataRow[k] = cell.StringCellValue;
                                        break;
                                    case CellType.Boolean:
                                        dataRow[k] = cell.BooleanCellValue;
                                        break;
                                    default:
                                        dataRow[k] = string.Empty;
                                        break;
                                }
                            }
                        }

                        table.Rows.Add(dataRow);
                    }
                }

                return table;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts Excel to DataTable.
        /// </summary>
        /// <param name="stream">The memory stream to be converted.</param>
        /// <param name="contentType">The mime content type.</param>
        /// <param name="isHeader">Whether the first row contains headers or not.</param>
        /// <returns></returns>
        public static DataTable Excel2DataTable(MemoryStream stream, string contentType, bool isHeader = true, int sheetNo = 0)
        {
            DataTable table = null;
            IWorkbook workbook = null;
            int startRow = isHeader ? 1 : 0;

            try
            {
                // 2007
                if (contentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    workbook = new XSSFWorkbook(stream);
                }
                // 2003
                else if (contentType == "application/vnd.ms-excel")
                {
                    workbook = new HSSFWorkbook(stream);
                }
                else
                {
                    throw new ArgumentException("The type of imported file is invalid. This function only accepts Excels(xls/xlsx).", "fileName");
                }

                if (workbook != null)
                {
                    ISheet sheet = workbook.GetSheetAt(sheetNo);
                    if (sheet == null)
                    {
                        throw new NullReferenceException("This sheet is null.");
                    }

                    IRow row = null;
                    ICell cell = null;

                    table = new DataTable(sheet.SheetName);

                    if (isHeader)
                    {
                        row = sheet.GetRow(0);
                        for (int k = 0; k < row.LastCellNum; k++)
                        {
                            cell = row.GetCell(k);
                            table.Columns.Add(cell.StringCellValue);
                        }
                    }

                    for (int j = startRow; j <= sheet.LastRowNum; ++j)
                    {
                        row = sheet.GetRow(j);
                        if (row == null)
                        {
                            continue;
                        }

                        DataRow dataRow = table.NewRow();

                        for (int k = 0; k < row.LastCellNum; k++)
                        {
                            cell = row.GetCell(k);
                            if (cell == null)
                            {
                                dataRow[k] = string.Empty;
                            }
                            else
                            {
                                // CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        dataRow[k] = cell.NumericCellValue;
                                        break;
                                    case CellType.String:
                                        dataRow[k] = cell.StringCellValue;
                                        break;
                                    case CellType.Formula:
                                        dataRow[k] = cell.StringCellValue;
                                        break;
                                    case CellType.Boolean:
                                        dataRow[k] = cell.BooleanCellValue;
                                        break;
                                    default:
                                        dataRow[k] = string.Empty;
                                        break;
                                }
                            }
                        }

                        table.Rows.Add(dataRow);
                    }

                }

                return table;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
