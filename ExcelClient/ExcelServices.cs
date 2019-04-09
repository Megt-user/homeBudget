using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Drawing;
using System.Globalization;
using System.IO;
using Transactions.Models;
using Transactions.Services;
using System;

namespace ExcelClient
{
    public class ExcelServices
    {
        private static int _startRow = 3;
        private static int _startColumn = 1;
        private static string _heding = "B1";
        private static string _titlecell = "C1";



        public static void CreateSheetWithTransactionMovments(List<MovementsViewModel> modementsViewModels, ExcelPackage excelPkg, string sheetName, string sheetHeading, string tableName)
        {
            ExcelWorksheet wsSheet;
            if (string.IsNullOrEmpty(sheetName))
                wsSheet = excelPkg.Workbook.Worksheets.Add("sheet01");
            else
                wsSheet = excelPkg.Workbook.Worksheets.Add(sheetName);

            //Add Table Title
            AddSheetHeading(wsSheet, sheetHeading);

            //Add transactions to excel Sheet
            CreateExcelTableFromMovementsViewModel(modementsViewModels, wsSheet, tableName);
        }

        public static void CreateSheetWithMonthSummary(List<MovementsViewModel> modementsViewModels, ExcelPackage excelPkg, string sheetName, string sheetHeading, IEnumerable<string> categoryList, string tableStartAdress = null)
        {

            if (!string.IsNullOrEmpty(tableStartAdress))
                SetStartRowAndColum(tableStartAdress);

            ExcelWorksheet wsSheet;
            if (string.IsNullOrEmpty(sheetName))
                wsSheet = excelPkg.Workbook.Worksheets.Add("MonthSummaries");
            else
                wsSheet = excelPkg.Workbook.Worksheets.Add(sheetName);

            //Add Table Title
            AddSheetHeading(wsSheet, "TableName");

            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, wsSheet, categoryList);
        }
        public static void CreateYearExpensesTable(List<MovementsViewModel> modementsViewModels, IEnumerable<string> categoryList, int year,
            ExcelWorksheet workSheet, string tableName, string tableStartAdress)
        {

            SetStartRowAndColum(tableStartAdress);
            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, workSheet, categoryList, year, "YearExpenses");
        }



        public static ExcelWorksheet GetExcelWorksheet(Stream streamFile, string sheetName = null)
        {
            ExcelPackage ep = new ExcelPackage(streamFile);
            ExcelWorksheet workSheet;
            if (string.IsNullOrEmpty(sheetName))
                workSheet = ep.Workbook.Worksheets.FirstOrDefault();
            else
                workSheet = ep.Workbook.Worksheets[sheetName];

            return workSheet;
        }

        public string GetJson(ExcelWorksheet ws)
        {
            var table = ws.Tables.FirstOrDefault();
            string json = string.Empty;
            if (table != null)
            {
                var TableStartRow = table.Address.Start.Row;
                var TableEndtRow = table.Address.End.Row;
                var dictionaryList = new List<Dictionary<string, string>>();

                for (int i = TableStartRow + 1; i <= TableEndtRow; i++)
                {
                    var valuesDictionary = new Dictionary<string, string>();
                    for (int j = table.Address.Start.Column; j <= table.Address.End.Column; j++)
                    {
                        //var headingPosition = string.Concat(GetColumnName(j), TableStartRow);
                        var cellTitle = ws.Cells[string.Concat(GetColumnName(j), TableStartRow)].Value;
                        //var cellName = string.Concat(GetColumnName(j), i);
                        var cellValue = ws.Cells[string.Concat(GetColumnName(j), i)].Value;
                        if (j == table.Address.Start.Column)
                            valuesDictionary.Add("Id", i.ToString());

                        if (valuesDictionary.ContainsKey(cellTitle.ToString()))
                        {
                            valuesDictionary.Add($"{cellTitle}_{string.Concat(GetColumnName(j), i)}", cellValue?.ToString());
                        }
                        else
                        {
                            valuesDictionary.Add(cellTitle.ToString(), cellValue?.ToString());
                        }
                    }
                    dictionaryList.Add(valuesDictionary);
                }
                json = Newtonsoft.Json.JsonConvert.SerializeObject(dictionaryList);
            }
            return json;
        }



        public static void AddSheetHeading(ExcelWorksheet wsSheet, string tableTitle)
        {
            wsSheet.Cells[_heding].Value = "Table Name";
            wsSheet.Cells[_titlecell].Value = tableTitle;
            wsSheet.Cells[_heding].Style.Font.Size = 12;
            wsSheet.Cells[_heding].Style.Font.Bold = true;
            wsSheet.Cells[_heding].Style.Font.Italic = true;
        }
        public static void AddTableHeadings(ExcelWorksheet wsSheet, string[] columnsNames, int rows)
        {
            using (ExcelRange rng = wsSheet.Cells[_startRow, _startColumn, _startRow, columnsNames.Length - 1])
            {
                tableHeadingFormat(rng, "Input");
            }
        }
        public static void CreateExcelMonthSummaryTableFromMovementsViewModel(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, IEnumerable<string> excelColumns, int sheetYear = 0, string sheetTableName = null)
        {
            int minYear;
            int maxYear;
            IEnumerable<string> TemExcelColumn;
            if (sheetYear > 0)
            {
                minYear = sheetYear;
                maxYear = sheetYear;
                //add month ant total column to Ienumeration
                TemExcelColumn = new[] { "Month", "Total" };
            }
            else
            {
                minYear = movementsModel.Min(mov => mov.DateTime.Year);
                maxYear = movementsModel.Max(mov => mov.DateTime.Year);
                //add month column to Ienumeration
                TemExcelColumn = new[] { "Month" };
            }

            // and the new columns to the category
            var newExcelColumn = TemExcelColumn.Concat(excelColumns);


            // Calculate size of the table
            var endRow = _startRow + 12;
            var endColum = _startColumn + newExcelColumn.Count();

            // Create Excel table Header
            int startRow = _startRow;
            int startColumn = _startColumn;

            //var row = _startRow + 1;
            var row = _startRow;
            for (int year = minYear; year <= maxYear; year++)
            {
                //give table Name
             var   tableName = sheetTableName ?? string.Concat("Table-", year);

                // add Headers to table
                CreateExcelTableHeader(wsSheet, tableName, newExcelColumn, startRow, endRow, _startColumn, endColum, true);

                var tableStartColumn = _startColumn;
                row++;

                // Set Excel table content
                for (int month = 1; month <= 12; month++)
                {
                    var monthName = string.Concat(DateTimeFormatInfo.CurrentInfo.GetMonthName(month));

                    AddExcelCellValue(tableStartColumn, row, monthName, wsSheet);
                    if (sheetYear > 0)
                    {
                        //Get summ for category
                        double? totalCategory = ModelClassServices.GetCategoriesMonthYearTotal(movementsModel, year, month);
                        AddExcelCellValue(tableStartColumn + 1, row, totalCategory, wsSheet);
                    }
                    foreach (var category in newExcelColumn)
                    {
                        if (category != "Month" && category != "Total")
                        {
                            //Get summ for category
                            double? totalCategory = ModelClassServices.GetTotalforCategory(movementsModel, category, year, month);

                            //add value tu excel cell
                            AddExcelCellValue(tableStartColumn, row, totalCategory, wsSheet);
                        }
                        tableStartColumn++;
                    }
                    tableStartColumn = _startColumn;
                    row++;
                }
                row = row + 2;
                startRow = row;
                endRow = row + 12;
            }

            //var noko = dict2.Keys.Except(dict.Keys);
            //var noko2 = dict.Keys.Except(dict2.Keys);
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        private static void CreateExcelTableHeader(ExcelWorksheet wsSheet, string tableName, IEnumerable<string> excelColumns, int startRow, int endRow, int startColumn, int endColum, bool ShowTotal = false)
        {
            using (ExcelRange rng = wsSheet.Cells[startRow, startColumn, endRow, endColum - 1])
            {
                //Indirectly access ExcelTableCollection class
                ExcelTable table = wsSheet.Tables.Add(rng, tableName);
                var color = Color.FromArgb(250, 199, 111);
                //Set Columns position & name
                var i = 0;
                foreach (var property in excelColumns)
                {
                    table.Columns[i].Name = string.Concat(property);
                    // add aggregate formulas (get these from an existing table in Excel)
                    //table.Columns[1].TotalsRowFormula = "SUBTOTAL(103,[Column1])"; // count

                    if (i != 0)
                        table.Columns[i].TotalsRowFormula = $"SUBTOTAL(109,[{property}])"; // sum 
                    //table.Columns[3].TotalsRowFormula = "SUBTOTAL(101,[Column3])"; // average
                    i++;
                }

                // Add empty cell for annotation
                //table.Columns[i].Name = "Annotation";
                //AddExcelCellByRowAndColumn(stratColumn + i, stratRow + 1, " ", wsSheet, color);

                table.ShowHeader = true;
                table.ShowFilter = true;
                table.ShowTotal = ShowTotal;
            }
        }

        public static void CreateExcelTableFromMovementsViewModel(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, string tableName)
        {
            var excelColumns = MovementsViewModel.excelColumns;

            // Calculate size of the table
            var endRow = _startRow + movementsModel.Count + 1;
            var endColum = _startColumn + excelColumns.Count;// TODO delete the last column in teblae

            // Create Excel table Header
            CreateExcelTableHeader(wsSheet, tableName, excelColumns, _startRow, endRow, _startColumn, endColum);



            // Set Excel table content
            var tableStartColumn = _startColumn;
            var row = _startRow + 1;
            foreach (var movement in movementsModel)
            {
                // account Movements
                foreach (var propertyName in excelColumns)
                {
                    //Get Property name value
                    var propertyValue = ModelClassServices.GetPropertyValue(movement, propertyName);
                    //add value tu excel cell
                    AddExcelCellValue(tableStartColumn, row, propertyValue, wsSheet);
                    tableStartColumn++;
                }
                tableStartColumn = _startColumn;
                row++;
            }

            //var noko = dict2.Keys.Except(dict.Keys);
            //var noko2 = dict.Keys.Except(dict2.Keys);
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        public static void UpdateTableValues(List<MovementsViewModel> movementsViewModels, Dictionary<string, string> categoriesAddressdictionary, int year, 
            ExcelWorksheet workSheet, string tableName, string columnName, string dictionaryKey)
        {
            //TODO check if dictionary have key and if Table have column name

            var addressDictionary = GetTableStartAdress(workSheet, tableName);

            // Get cell address
            var columnNameAdress = GetColumnNameAdress(columnName, workSheet, tableName);
            var dictionaryKeyAddress = categoriesAddressdictionary[dictionaryKey];

            //Get Row and Colum Index
            var columNamecellIndex = GetRowAndColumIndex(columnNameAdress);

            if (addressDictionary.Any())
            {
                for (int month = 1; month <= 12; month++)
                {
                    string newCellAdress = AddRowAndColumnToCellAddress(categoriesAddressdictionary[dictionaryKey], month, 0);
                    AddExcelCellFormula(columNamecellIndex["column"], columNamecellIndex["row"] + month, newCellAdress, workSheet);
                }

            }
        }

        private static string AddRowAndColumnToCellAddress(string address, int row, int column)
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

        private static SubCategory SetSubCategoryToExpense(IEnumerable<SubCategory> subcategoriesMatch)
        {
            var subcategory = new SubCategory();
            var moreThanOneCategory = subcategoriesMatch.Select(sub => sub.Category).Distinct().Count() > 1;
            if (moreThanOneCategory)
            {
                var subcategoryNames = subcategoriesMatch.Select(sub => sub.KeyWord).ToArray();
                var subcategoryCategories = subcategoriesMatch.Select(sub => sub.Category).ToArray();
                subcategory.KeyWord = string.Join(",", subcategoryNames);
                subcategory.Category = string.Join(",", subcategoryCategories);
                subcategory.SupPorject = "Mismatch";
            }
            else
            {
                subcategory = subcategoriesMatch.FirstOrDefault();
            }
            return subcategory;
        }

        private static void AddSubcategoryToExcel(ExcelWorksheet wsSheet, List<string> subCategoryProperties, int subCategoryColumn, int row, SubCategory subcategory)
        {
            foreach (var property in subCategoryProperties)
            {
                string valueString = valueString = subcategory.GetType().GetProperty(property).GetValue(subcategory, null)?.ToString();
                AddExcelCellValue(subCategoryColumn, row, valueString, wsSheet);
                subCategoryColumn++;
            }
        }

        // internal method
        private static void tableHeadingFormat(ExcelRange excelRange, string text)
        {
            excelRange.Value = text;
            excelRange.Style.Font.Size = 12;
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Italic = true;
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelRange.Merge = true;
            excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }
        public static void AddExcelCellValue(int column, int row, object value, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellName = string.Concat(GetColumnName(column), row);
            using (ExcelRange rng1 = wsSheet.Cells[cellName])
            {
                if (color.HasValue)
                {
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(color.Value);
                }
                rng1.Value = value;
            }
        }
        public static void AddExcelCellFormula(int column, int row, object formula, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellName = string.Concat(GetColumnName(column), row);
            using (ExcelRange rng1 = wsSheet.Cells[cellName])
            {
                if (color.HasValue)
                {
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(color.Value);
                }
                rng1.Formula = $"={formula}";
            }
        }
        public static void AddExcelCellValue(string cellAddress, object value, ExcelWorksheet wsSheet, Color? color = null)
        {
            using (ExcelRange rng1 = wsSheet.Cells[cellAddress])
            {
                if (color.HasValue)
                {
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(color.Value);
                }
                rng1.Value = value;
            }
        }

        public static int GetColumnIndex(string columnName)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var index = letters.ToLower().IndexOf(columnName.ToLower());

            return index + 1;
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


        public static void SetStartRowAndColum(string cellAdress)
        {
            if (!string.IsNullOrEmpty(cellAdress))
            {
                var column = string.Empty;
                var row = string.Empty;
                foreach (char c in cellAdress)
                {
                    if (char.IsLetter(c))
                        column += c;
                    if (char.IsNumber(c))
                        row += c;
                }
                int rowNumber;
                int.TryParse(row, out rowNumber);

                _startRow = rowNumber;
                _startColumn = GetColumnIndex(column);
            }
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
                dictionay.Add("column", ExcelServices.GetColumnIndex(column));
                if (addressAndWorkSheet.Length > 1)
                {
                    dictionay.Add("WorkSheet", 0);
                }
                return dictionay;
            }
            return null;
        }

        public static Dictionary<string, string> GetColumnsNameAdress(IEnumerable<string> categories, ExcelWorksheet workSheet, string tableName, int row = 0)
        {
            var dictionary = new Dictionary<string, string>();



            var addressDictionary = GetTableStartAdress(workSheet, tableName);

            if (addressDictionary.Any())
            {
                foreach (var item in categories)
                {
                    int idx = GetIndexFromColumnName(workSheet, addressDictionary["row"], item);

                    if (idx > 0)
                    {
                        dictionary.Add(item, $"'{workSheet.Name}'!{GetColumnName(idx)}{addressDictionary["row"]}");
                    }
                }
            }
            return dictionary;
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

        private static Dictionary<string, int> GetTableStartAdress(ExcelWorksheet workSheet, string tableName)
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


    }



}

