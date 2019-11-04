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
        private static int _startRow;
        private static int _startColumn;
        private static string _heding;
        private static string _titlecell;

        public ExcelServices()
        {
            _startRow = 3;
            _startColumn = 1;
            _heding = "B1";
            _titlecell = "C1";
        }

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

        public static void CreateSheetWithMonthSummary(List<MovementsViewModel> modementsViewModels, ExcelPackage excelPkg, string sheetName, IEnumerable<string> categoryList)
        {

            ExcelWorksheet wsSheet;
            if (string.IsNullOrEmpty(sheetName))
                wsSheet = excelPkg.Workbook.Worksheets.Add("MonthSummaries");
            else
                wsSheet = excelPkg.Workbook.Worksheets.Add(sheetName);

            //Add Table Title
            AddSheetHeading(wsSheet, "TableName");

            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, wsSheet, categoryList, 0, null, true);
        }
        public static void CreateYearExpensesTable(List<MovementsViewModel> modementsViewModels, IEnumerable<string> categoryList, int year,
            ExcelWorksheet workSheet, string tableName, string tableStartAdress)
        {

            SetStartRowAndColum(tableStartAdress);
            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, workSheet, categoryList, year, tableName, true);
        }
        public static void CreateYearIncomsTable(List<MovementsViewModel> modementsViewModels, IEnumerable<string> categoryList, int year,
            ExcelWorksheet workSheet, string tableName, string tableStartAdress)
        {

            SetStartRowAndColum(tableStartAdress);
            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, workSheet, categoryList, year, tableName, false);
        }
        public static void CreateCategoriesAverageTable(List<MovementsViewModel> modementsViewModels, IEnumerable<string> categoryList, int year, int month,
            ExcelWorksheet workSheet, string tableName)
        {

            //Add transactions to excel Sheet
            CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, workSheet, categoryList, year, tableName, false);
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

        public static string GetJsonFromTable(ExcelWorksheet ws, string tableName = null)
        {
            ExcelTable table;
            if (string.IsNullOrEmpty(tableName))
            {
                table = ws.Tables.FirstOrDefault();
            }
            else
            {
                table = ws.Tables[tableName];
            }

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
                        var cellTitle = ws.Cells[string.Concat(ExcelHelpers.GetColumnName(j), TableStartRow)].Value;
                        //var cellName = string.Concat(GetColumnName(j), i);
                        var cellValue = ws.Cells[string.Concat(ExcelHelpers.GetColumnName(j), i)].Value;
                        if (j == table.Address.Start.Column)
                            valuesDictionary.Add("Id", i.ToString());

                        if (valuesDictionary.ContainsKey(cellTitle.ToString()))
                        {
                            valuesDictionary.Add($"{cellTitle}_{string.Concat(ExcelHelpers.GetColumnName(j), i)}", cellValue?.ToString());
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

        public static void CreateCategoriesAverage(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, IEnumerable<string> categories,
            int year = 0, int month1 = 0, bool justExtrations = true)
        {
            if (justExtrations)
            {
                categories = ModelClassServices.GetExtractionCategories(categories, movementsModel);
            }
            else
            {
                categories = ModelClassServices.GetIncomsCategories(categories, movementsModel);
            }
            categories = categories.OrderBy(c => c);
            var columns = new[] { "Type" };
            var newAverageColumn = new[] { $"Year({year})", "Day" };

            var newExcelColumn = columns.Concat(categories);
            // Calculate size of the table
            var endRow = _startRow + newAverageColumn.Count();
            var endColum = _startColumn + newExcelColumn.Count();

            // Create Excel table Header
            int startRow = _startRow;
            int startColumn = _startColumn;

            var row = _startRow;

            var tableStartColumn = _startColumn;

            // Set Excel table content

            //give table Name
            var tableName = "CategoriesAverage";
            // Add table Headers
            CreateExcelTableHeader(wsSheet, tableName, newExcelColumn, startRow, endRow, _startColumn, endColum, true);
            row++;

            foreach (var item in newAverageColumn)
            {
                foreach (var category in newExcelColumn)
                {
                    if (category == "Type")
                    {
                        AddExcelCellValue(row, tableStartColumn, item, wsSheet);
                    }
                    else
                    {
                        double categoryAverage = 0;
                        if (item == $"Year({year})")
                        {
                            categoryAverage = ModelClassServices.AverageforCategory(movementsModel, category, year, null, justExtrations);
                        }
                        if (item == $"Day")
                        {
                            categoryAverage = ModelClassServices.AverageforCategory(movementsModel, category, null, null, justExtrations);
                        }
                        wsSheet.Cells[row, tableStartColumn].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                        AddExcelCellValue(row, tableStartColumn, categoryAverage, wsSheet);
                    }
                    tableStartColumn++;
                }
                tableStartColumn = _startColumn;
                row++;
            }

            row = row + 3;

            // Calculate size of the table
            endRow = row + 12;
            endColum = _startColumn + newExcelColumn.Count();

            // Create Excel table Header
            startRow = row;
            startColumn = _startColumn;


            tableStartColumn = _startColumn;

            // Set Excel table content

            //give table Name
            tableName = "CategoriesMonthAverage";
            // Add table Headers
            CreateExcelTableHeader(wsSheet, tableName, newExcelColumn, startRow, endRow, _startColumn, endColum, true);
            row++;

            tableStartColumn = _startColumn;

            for (int month = 1; month <= 12; month++)
            {
                var monthName = string.Concat(DateTimeFormatInfo.CurrentInfo.GetMonthName(month));
                AddExcelCellValue(row, tableStartColumn, monthName, wsSheet);

                foreach (var category in newExcelColumn)
                {
                    double categoryAverage = 0;

                    if (category == "Type")
                    {
                        tableStartColumn++;
                        continue;
                    }
                    else
                    {
                        categoryAverage = ModelClassServices.AverageforCategory(movementsModel, category, null, month, justExtrations);

                        wsSheet.Cells[row, tableStartColumn].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                        AddExcelCellValue(row, tableStartColumn, categoryAverage, wsSheet);
                        tableStartColumn++;
                    }
                }
                tableStartColumn = _startColumn;
                row++;
            }

        }
        public static void CreateExcelMonthSummaryTableFromMovementsViewModel(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, IEnumerable<string> categories,
            int sheetYear = 0, string sheetTableName = null, bool justExtrations = true)
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

            if (justExtrations)
            {
                categories = ModelClassServices.GetExtractionCategories(categories, movementsModel);
            }
            else
            {
                categories = ModelClassServices.GetIncomsCategories(categories, movementsModel);
            }


            categories = categories.OrderBy(c => c);
            // and the new columns to the category
            var newExcelColumn = TemExcelColumn.Concat(categories);


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
                var tableName = sheetTableName ?? string.Concat("Table-", year);

                // add Headers to table
                CreateExcelTableHeader(wsSheet, tableName, newExcelColumn, startRow, endRow, _startColumn, endColum, true);

                var tableStartColumn = _startColumn;
                row++;

                // Set Excel table content
                for (int month = 1; month <= 12; month++)
                {
                    var monthName = string.Concat(DateTimeFormatInfo.CurrentInfo.GetMonthName(month));
                    AddExcelCellValue(row, tableStartColumn, monthName, wsSheet);
                    if (sheetYear > 0)
                    {
                        //Get summ for category
                        double totalCategory = ModelClassServices.CategoriesMonthYearTotal(movementsModel, year, month, justExtrations);

                        AddExcelCellValue(row, tableStartColumn + 1, totalCategory, wsSheet);
                    }
                    foreach (var category in newExcelColumn)
                    {
                        if (category != "Month" && category != "Total")
                        {
                            //Get summ for category
                            double totalCategory = ModelClassServices.TotalforCategory(movementsModel, category, year, month, justExtrations);
                            wsSheet.Cells[row, tableStartColumn].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                            //add value tu excel cell
                            AddExcelCellValue(row, tableStartColumn, totalCategory, wsSheet);
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

                //var color = Color.FromArgb(250, 199, 111);

                //Set Columns position & name
                var i = 0;
                foreach (var property in excelColumns)
                {
                    table.Columns[i].Name = string.Concat(property);

                    //Add total cell to the end of the table
                    if (i != 0)
                        table.Columns[i].TotalsRowFormula = $"SUBTOTAL(101,[{property}])"; // 101 average, 103 Count, 109 sum

                    i++;
                }

                table.ShowHeader = true;
                table.ShowFilter = true;
                table.ShowTotal = ShowTotal;
            }
        }

        public static void CreateExcelTableFromMovementsViewModel(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, string tableName)
        {
            //Get the list of Column that want to be created in the table
            var excelColumns = MovementsViewModel.excelColumns;

            // Calculate size of the table
            var endRow = _startRow + movementsModel.Count + 1;
            var endColum = _startColumn + excelColumns.Count;

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
                    wsSheet.Cells[row, tableStartColumn].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell(propertyName);

                    AddExcelCellValue(row, tableStartColumn, propertyValue, wsSheet);
                    tableStartColumn++;
                }
                tableStartColumn = _startColumn;
                row++;
            }

            //var noko = dict2.Keys.Except(dict.Keys);
            //var noko2 = dict.Keys.Except(dict2.Keys);
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        public static void UpdateYearTableValues(Dictionary<string, string> categoriesAddressdictionary, int year,
            ExcelWorksheet workSheet, string tableName, string columnName, string dictionaryKey)
        {
            //TODO check if dictionary have key and if Table have column name

            var addressDictionary = ExcelHelpers.GetTableStartAdress(workSheet, tableName);

            // Get cell address
            var columnNameAdress = ExcelHelpers.GetColumnNameAdress(columnName, workSheet, tableName);
            var dictionaryKeyAddress = categoriesAddressdictionary[dictionaryKey];

            //Get Row and Colum Index
            var columNamecellIndex = ExcelHelpers.GetRowAndColumIndex(columnNameAdress);

            if (addressDictionary.Any())
            {
                for (int month = 1; month <= 12; month++)
                {
                    string newCellAdress = ExcelHelpers.AddRowAndColumnToCellAddress(categoriesAddressdictionary[dictionaryKey], month, 0);
                    workSheet.Cells[columNamecellIndex["row"], columNamecellIndex["column"]].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                    AddExcelCellFormula(columNamecellIndex["row"] + month, columNamecellIndex["column"], newCellAdress, workSheet);
                }

            }
        }
        public static void UpdateClassesTableValues(Dictionary<string, string> BudgetCategoriesAddressdictionary, Dictionary<string, string> ExpenseCategoriesAddressdictionary, int year,
            ExcelWorksheet workSheet, string tableName)
        {
            //TODO check if dictionary have key and if Table have column name

            var addressDictionary = ExcelHelpers.GetTableStartAdress(workSheet, tableName);

            //TODO check cell Value
            var date = (double)workSheet.Cells["G1"].Value;
            var monthToFilter = DateTime.FromOADate(date).Month;

            // Get cell address
            var OperatingAdress = ExcelHelpers.GetColumnNameAdress("OPERATING", workSheet, tableName);
            var BudgetAdress = ExcelHelpers.GetColumnNameAdress("BUDGET", workSheet, tableName);
            var ActualAdress = ExcelHelpers.GetColumnNameAdress("ACTUAL", workSheet, tableName);

            //Get Row and Colum Index
            var OperatingIndex = ExcelHelpers.GetRowAndColumIndex(OperatingAdress);
            var BudgetIndex = ExcelHelpers.GetRowAndColumIndex(BudgetAdress);
            var ActualIndex = ExcelHelpers.GetRowAndColumIndex(ActualAdress);
            //int budgetCategories = BudgetCategoriesAddressdictionary.Count();
            //int expenseCategories = ExpenseCategoriesAddressdictionary.Count();

            var categories = BudgetCategoriesAddressdictionary.Where(ct => ExpenseCategoriesAddressdictionary.ContainsKey(ct.Key)).Select(ct => ct.Key).ToList();

            //List<string> categories = budgetCategories > expenseCategories ? new List<string>(ExpenseCategoriesAddressdictionary.Keys) : new List<string>(BudgetCategoriesAddressdictionary.Keys);

            if (addressDictionary.Any())
            {
                var i = 1;
                foreach (var category in categories)
                {
                    //TODO Update formula =HLOOKUP([@OPERATING];'Expenses details'!$E$22:$AC$34;MONTH($G$1)+1;FALSE)

                    string budgetCellAdress = BudgetCategoriesAddressdictionary[category];
                    string actualCellAdress = ExpenseCategoriesAddressdictionary[category];

                    string newBudgetCell = $"OFFSET({budgetCellAdress},MONTH($G$1),0)";
                    string newActualCell = $"OFFSET({actualCellAdress},MONTH($G$1),0)";

                    AddExcelCellValue(OperatingIndex["row"] + i, OperatingIndex["column"], category, workSheet);

                    workSheet.Cells[BudgetIndex["row"] + i, BudgetIndex["column"]].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                    AddExcelCellFormula(BudgetIndex["row"] + i, BudgetIndex["column"], newBudgetCell, workSheet);
                    workSheet.Cells[ActualIndex["row"] + i, ActualIndex["column"]].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                    AddExcelCellFormula(ActualIndex["row"] + i, ActualIndex["column"], newActualCell, workSheet);
                    i++;
                }
            }
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
                AddExcelCellValue(row, subCategoryColumn, valueString, wsSheet);
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
        public static void AddExcelCellFormula(int row, int column, string formula, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellAddress = $"{ExcelHelpers.GetColumnName(column)}{row}";
            AddExcelCellFormula(cellAddress, formula, wsSheet, color);
        }

        public static void AddExcelCellFormula(string cellAddress, string formula, ExcelWorksheet wsSheet, Color? color = null)
        {
            wsSheet.Cells[cellAddress].Formula = formula;
        }

        public static void AddExcelCellValue(string cellAddress, object value, ExcelWorksheet wsSheet, Color? color = null)
        {
            if (color.HasValue)
            {
                wsSheet.Cells[cellAddress].Style.Fill.PatternType = ExcelFillStyle.Solid;
                wsSheet.Cells[cellAddress].Style.Fill.BackgroundColor.SetColor(color.Value);
            }
            wsSheet.Cells[cellAddress].Value = value;
        }
        public static void AddExcelCellValue(int row, int column, object value, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellAddress = $"{ExcelHelpers.GetColumnName(column)}{row}";

            AddExcelCellValue(cellAddress, value, wsSheet, color);

        }

        public static Dictionary<string, string> GetColumnsNameAdress(IEnumerable<string> categories, ExcelWorksheet workSheet, string tableName, int row = 0)
        {
            var dictionary = new Dictionary<string, string>();

            var addressDictionary = ExcelHelpers.GetTableStartAdress(workSheet, tableName);

            if (addressDictionary.Any())
            {
                foreach (var item in categories)
                {
                    int idx = ExcelHelpers.GetIndexFromColumnName(workSheet, addressDictionary["row"], item);

                    if (idx > 0)
                    {
                        dictionary.Add(item, $"'{workSheet.Name}'!{ExcelHelpers.GetColumnName(idx)}{addressDictionary["row"]}");
                    }
                }
            }
            return dictionary;
        }
        private static void SetStartRowAndColum(string cellAdress)
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
                _startColumn = ExcelHelpers.GetColumnIndex(column);
            }
        }

    }


}

