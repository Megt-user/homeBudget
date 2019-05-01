using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelClient
{
   public class ExcelHelpers
    {
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

        public static Dictionary<string, int> GetRowAndColumIndex(string address)
        {

            if (!string.IsNullOrEmpty(address))
            {

                var addressAndWorkSheet = address.Split("!");

                var cellAddress = addressAndWorkSheet.Length > 1 ? addressAndWorkSheet[1] : addressAndWorkSheet[0];


                Dictionary<string, int> dictionay = new Dictionary<string, int>();

                var column = string.Empty;
                var row = string.Empty;

                foreach (char c in cellAddress)
                {
                    if (char.IsLetter(c))
                        column += c;
                    if (char.IsNumber(c))
                        row += c;
                }
                int rowNumber;
                int.TryParse(row, out rowNumber);

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
            if (!string.IsNullOrEmpty(tableName) && row == 0)
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
            if (!string.IsNullOrEmpty(columnName) && row > 0 && workSheet != null)
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

        public static string SetFormatToCell(string value)
        {
            switch (value)
            {
                case "Id":
                    return "#";
                case "DateTime":
                    return "dd/mm/yyyy";
                case "Amount":
                    return "$ # ##0.00;[Red]$ -# ##0.00";
                default:
                    return "#";
            }
        }

    }
}
