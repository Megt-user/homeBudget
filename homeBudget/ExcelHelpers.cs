using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace homeBudget
{
    public class ExcelHelpers
    {
        /// <summary>
        /// Checked
        /// </summary>
        /// <param name="address"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static string AddRowAndColumnToCellAddress(string address, int row, int column)
        {
            var addressAndWorkSheet = address.Split("!");

            var cellAddress = addressAndWorkSheet.Length > 1 ? addressAndWorkSheet[1] : addressAndWorkSheet[0];

            var dictionaryKeyIndex = GetRowAndColumIndex(cellAddress);

            if (dictionaryKeyIndex.Any())
            {
                var newaddress = $"{GetColumnName(dictionaryKeyIndex["column"] + column)}{dictionaryKeyIndex["row"] + row}";
                return addressAndWorkSheet.Length > 1 ? $"{addressAndWorkSheet[0]}!{newaddress}" : newaddress;
            }
            return null;
        }
        /// <summary>
        /// Checked
        /// </summary>
        /// <param name="categories"></param>
        /// <param name="excelTable"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetNamesAdress(IEnumerable<string> categories, ExcelTable excelTable)
        {
            var dictionary = new Dictionary<string, string>();
            foreach (var item in categories)
            {
                var categoryColumnIndex = ExcelHelpers.GetAdressFromColumnName(excelTable, item);
                if (!String.IsNullOrEmpty(categoryColumnIndex))
                    dictionary.Add(item, categoryColumnIndex);

            }
            return dictionary;
        }
        /// <summary>
        /// Checked
        /// </summary>
        /// <param name="excelTable"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string GetAdressFromColumnName(ExcelTable excelTable, string columnName)
        {
            string cellAdress = null;
            var stratRow = excelTable.Address.Start.Address;
            var endRow = AddRowAndColumnToCellAddress(excelTable.Address.Start.Address, 0, excelTable.Address.Columns);
            var valueExist = excelTable.WorkSheet.Cells[$"{stratRow}:{endRow}"].Any(c => c.Value?.ToString().ToLower() == columnName.ToLower());
            if (valueExist)
            {
                cellAdress = excelTable.WorkSheet
                    .Cells[$"{stratRow}:{endRow}"]
                    .First(c => String.Equals(c.Value?.ToString(), columnName, StringComparison.CurrentCultureIgnoreCase))
                    .Address;
            }
            return cellAdress != null ? $"'{excelTable.WorkSheet.Name}'!{cellAdress}" : null;
        }
        public static Dictionary<string, int> GetRowAndColumIndex(string address)
        {

            if (!String.IsNullOrEmpty(address))
            {

                var addressAndWorkSheet = address.Split("!");

                var cellAddress = addressAndWorkSheet.Length > 1 ? addressAndWorkSheet[1] : addressAndWorkSheet[0];


                Dictionary<string, int> dictionay = new Dictionary<string, int>();

                var column = String.Empty;
                var row = String.Empty;

                foreach (char c in cellAddress)
                {
                    if (Char.IsLetter(c))
                        column += c;
                    if (Char.IsNumber(c))
                        row += c;
                }
                int rowNumber;
                Int32.TryParse(row, out rowNumber);

                dictionay.Add("row", rowNumber);
                dictionay.Add("column", GetColumnIndex(column));
                if (addressAndWorkSheet.Length > 1)
                {
                    dictionay.Add("WorkSheet", 0);
                }
                return dictionay;
            }
            return null;
        }
        public static int GetColumnIndex(string columnName)
        {
            var index = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                index *= 26;
                index += (columnName[i] - 'A' + 1);
            }

            return index;
        }

        public static string GetColumnName(int index)
        {
            int dividend = index;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        //TODO Check This methods

        public static Dictionary<string, int> GetTableStartAdress(ExcelWorksheet workSheet, string tableName)
        {
            var addressDictionary = new Dictionary<string, int>();
            var exTable = workSheet.Tables[tableName];
            if (exTable != null)
            {
                var tableStartAdress = exTable.Address.Start.Address;
                addressDictionary = GetRowAndColumIndex(tableStartAdress);
            }
            return addressDictionary;
        }

        public static string GetColumnNameAdress(string columnName, ExcelWorksheet workSheet, string tableName, int row = 0)
        {
            if (!String.IsNullOrEmpty(tableName) && row == 0)
            {
                var addressDictionary = GetTableStartAdress(workSheet, tableName);
                if (addressDictionary.Any())
                {
                    row = addressDictionary["row"];
                }
            }

            if (row > 0)
            {
                int idx = GetIndexFromColumnName(workSheet, row, columnName);

                if (idx > 0)
                {
                    return $"'{workSheet.Name}'!{GetColumnName(idx)}{row}";
                }
            }
            return null;
        }
        public static string GetColumnNameCell(string columnName, ExcelWorksheet workSheet, string tableName, int row = 0)
        {
            if (!String.IsNullOrEmpty(tableName) && row == 0)
            {
                var addressDictionary = GetTableStartAdress(workSheet, tableName);
                if (addressDictionary.Any())
                {
                    row = addressDictionary["row"];
                }
            }

            if (row > 0)
            {
                int idx = GetIndexFromColumnName(workSheet, row, columnName);

                if (idx > 0)
                {
                    return $"'{workSheet.Name}'!{GetColumnName(idx)}{row}";
                }
            }
            return null;
        }
        public static int GetIndexFromColumnName(ExcelWorksheet workSheet, int row, string columnName)
        {
            if (!String.IsNullOrEmpty(columnName) && row > 0 && workSheet != null)
            {
                var valueExist = workSheet.Cells[$"{row}:{row}"].Any(c => c.Value?.ToString().ToLower() == columnName.ToLower());
                if (valueExist)
                {
                    return workSheet
                            .Cells[$"{row}:{row}"]
                            .First(c => c.Value?.ToString().ToLower() == columnName.ToLower())
                            .Start
                            .Column;
                }
            }
            return 0;
        }

        public static ExcelWorksheet GetExcelWorksheet(Stream streamFile, string sheetName = null)
        {
            ExcelPackage ep = new ExcelPackage(streamFile);
            ExcelWorksheet workSheet;
            if (String.IsNullOrEmpty(sheetName))
                workSheet = ep.Workbook.Worksheets.FirstOrDefault();
            else
                workSheet = ep.Workbook.Worksheets[sheetName];

            return workSheet;
        }

        public static ExcelTable GetWorkSheeTable(ExcelWorksheet excelWorksheet, string tableName = null)
        {
            return String.IsNullOrWhiteSpace(tableName) ? excelWorksheet.Tables.FirstOrDefault() : excelWorksheet.Tables[tableName];
        }

        //Read Excel Files
        public static async Task<ExcelWorksheet> GetExcelWorkSheet(IFormFile transacation, string filePathTemp)
        {
            ExcelWorksheet transactionsWorkSheet;
            using (var stream = new FileStream(filePathTemp, FileMode.Create))
            {
                await transacation.CopyToAsync(stream);
                transactionsWorkSheet = ExcelHelpers.GetExcelWorksheet(stream);
            }
            return transactionsWorkSheet;
        }

        public static string SetFormatToCell(string value)
        {
            switch (value)
            {
                case "Id":
                    return "#";
                case "DateTime":
                    return "dd/mm/yyyy";
                case "Amount":
                    return "$ #,##0.00;[Red]$ -#,##0.00";
                default:
                    return "#";
            }
        }
    }
}
