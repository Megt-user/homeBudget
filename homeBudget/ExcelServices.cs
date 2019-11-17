using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using homeBudget.Models;
using homeBudget.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace homeBudget
{
    public class ExcelServices
    {


        public static string _startCell = "A3";
        private static string _hedingCell = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, -2, 1);
        private static string _titlecell = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, -2, 2);


        /// <summary>
        /// 
        /// </summary>
        /// <param name="categories"></param>
        /// <param name="workSheet"></param>
        /// <param name="tableName"></param>
        /// <param name="row"></param>
        /// <param name="excelTable"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetColumnsNameAdress(IEnumerable<string> categories, ExcelTable excelTable)
        {
            var dictionary = new Dictionary<string, string>();
            foreach (var item in categories)
            {
                var cellAdress = ExcelHelpers.GetAdressFromColumnName(excelTable, item);
                if (!string.IsNullOrEmpty(cellAdress))
                {
                    dictionary.Add(item, cellAdress);
                }
            }
            return dictionary;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="movementsModel"></param>
        /// <param name="wsSheet"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static ExcelTable CreateExcelTableFromMovementsViewModel(List<TransactionViewModel> movementsModel, ExcelWorksheet wsSheet, string tableName)
        {
            //Get the list of Column that want to be created in the table
            var excelColumns = TransactionViewModel.excelColumns;

            // Calculate size of the table
            var endTableCellAdress = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, movementsModel.Count, excelColumns.Count - 1);

            // Create Excel table Header
            var excelTable = CreateExcelTable(wsSheet, tableName, excelColumns, _startCell, endTableCellAdress);

            for (int row = 0; row < movementsModel.Count; row++)
            {
                for (int column = 0; column < excelColumns.Count; column++)
                {
                    //Get Property name value
                    var propertyValue = ModelConverter.GetPropertyValue(movementsModel[row], excelColumns[column]);
                    string tableCellAdress = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, row + 1, column);
                    excelTable.WorkSheet.Cells[tableCellAdress].Value = propertyValue;
                }
            }
            excelTable.WorkSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
            return excelTable;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wsSheet"></param>
        /// <param name="excelPackage"></param>
        /// <param name="movementsModel"></param>
        /// <param name="categories"></param>
        /// <param name="sheetYear"></param>
        /// <param name="sheetTableName"></param>
        /// <param name="justExtrations"></param>
        /// <param name="startCell"></param>
        /// <returns></returns>
        public static List<ExcelTable> CreateExcelMonthSummaryTableFromMovementsViewModel(ExcelWorksheet wsSheet, List<TransactionViewModel> movementsModel, IEnumerable<string> categories,
            int sheetYear = 0, string sheetTableName = null, bool justExtrations = true, string startCell = null)
        {
           
            int minYear;
            int maxYear;
            if (sheetYear > 0)
            {
                minYear = sheetYear;
                maxYear = sheetYear;
            }
            else
            {
                minYear = movementsModel.Min(mov => mov.DateTime.Year);
                maxYear = movementsModel.Max(mov => mov.DateTime.Year);
            }

            IEnumerable<string> newColumns = new[] { "Month", "Total" };
            categories = justExtrations ? ModelOperation.GetExtractionCategories(categories, movementsModel) : ModelOperation.GetIncomsCategories(categories, movementsModel);
            categories = categories.OrderBy(c => c);

            // and the new columns to the category
            categories = newColumns.Concat(categories);
            var categoriesUpdated = categories as string[] ?? categories.ToArray();

            string startTableCell = startCell ?? _startCell;

            // Create Excel table Header
            var endTableCellAddress = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, 12, categoriesUpdated.Count() - 1);
            var tableName = sheetTableName ?? "Tanble-";
            List<ExcelTable> excelTables = new List<ExcelTable>();
            for (int year = minYear; year <= maxYear; year++)
            {
                //give table Name
                tableName = string.Concat(tableName, year);

                //calculate Table sizes
                endTableCellAddress = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, 12, categoriesUpdated.Count() - 1);
                var excelTable = CreateExcelTable(wsSheet, tableName, categoriesUpdated, startTableCell, endTableCellAddress, true);

                // Set Excel table content
                for (int month = 1; month <= 12; month++)
                {
                    for (int i = 0; i < categoriesUpdated.Length; i++)
                    {
                        switch (categoriesUpdated[i])
                        {
                            case "Month":
                                var monthName = string.Concat(DateTimeFormatInfo.CurrentInfo.GetMonthName(month));
                                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, month, i)].Value = monthName;
                                break;
                            case "Total":
                                double totalCategory = ModelConverter.CategoriesMonthYearTotal(movementsModel, year, month, justExtrations);
                                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, month, i)].Value = totalCategory;
                                break;
                            default:
                                //Get summ for category
                                var tablecell = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, month, i);
                                double totalCategory1 = ModelOperation.GetTotalforCategory(movementsModel, categoriesUpdated[i], year, month, justExtrations);
                                excelTable.WorkSheet.Cells[tablecell].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                                //add value tu excel cell
                                wsSheet.Cells[tablecell].Value = totalCategory1;
                                //AddExcelCellValue(row, tableStartColumn, totalCategory1, wsSheet);
                                break;
                        }
                    }
                }
                startTableCell = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, 12 + 5, 0);
                excelTable.WorkSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
                excelTables.Add(excelTable);
            }
            return excelTables;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="movementsModel"></param>
        /// <param name="wsSheet"></param>
        /// <param name="categories"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="justExtrations"></param>
        /// <returns></returns>
        public static ExcelTable CreateAverageForYearMonthDay(List<TransactionViewModel> movementsModel, ExcelWorksheet wsSheet, IEnumerable<string> categories,
            int year = 0, int? month = 0, bool justExtrations = true)
        {
            categories = justExtrations ? ModelOperation.GetExtractionCategories(categories, movementsModel) : ModelOperation.GetIncomsCategories(categories, movementsModel);

            categories = categories.OrderBy(c => c);
            var columns = new[] { "Type" };

            var newAverageColumn = new List<string>();

            //to get the month average we need to specify if is jus one year or all the years
            IEnumerable<int> years = null;

            if (year > 0)
            {
                newAverageColumn.Add($"Year({year})");
            }
            else
            {
                newAverageColumn.Add("Years");
                years = movementsModel.Select(mov => mov.DateTime.Year).Distinct();
                newAverageColumn.AddRange(years.Select(selectedYear => $"Year({selectedYear})"));
            }

            newAverageColumn.Add("Month");
            newAverageColumn.Add("Day");

            var newExcelColumn = columns.Concat(categories);
            // Calculate size of the table

            var excelColumns = newExcelColumn as string[] ?? newExcelColumn.ToArray();
            string endTableCell = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, newAverageColumn.Count(), excelColumns.Count() - 1);

            //give table Name
            var tableName = "CategoriesYearMonthDayAverage";
            // Add table Headers

            var excelTable = CreateExcelTable(wsSheet, tableName, excelColumns, _startCell, endTableCell, true);

            for (int row = 0; row < newAverageColumn.Count(); row++)
            {
                var rowValue = newAverageColumn[row];
                for (int column = 0; column < excelColumns.Count(); column++)
                {
                    var category = excelColumns[column];
                    string cellAdress = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, row + 1, column);
                    double categoryAverage = 0;

                    if (category == "Type")
                    {
                        excelTable.WorkSheet.Cells[cellAdress].Value = newAverageColumn[row];
                    }
                    else
                    {
                        if (rowValue == "Years")
                        {
                            categoryAverage = ModelOperation.AverageforCategory(movementsModel, category, null, null, justExtrations);
                        }
                        else if (rowValue == "Month")
                        {
                            categoryAverage = years != null ? ModelOperation.AverageforCategory(movementsModel, category, null, 0, justExtrations) :
                                ModelOperation.AverageforCategory(movementsModel, category, year, 0, justExtrations);
                        }
                        else
                        {

                            if (years != null)
                            {
                                foreach (var selectedYear in years)
                                {
                                    if (rowValue == $"Year({selectedYear})")
                                        categoryAverage = ModelOperation.GetTotalforCategory(movementsModel, category, selectedYear, null, justExtrations);
                                }
                            }
                            else
                            {
                                categoryAverage = ModelOperation.GetTotalforCategory(movementsModel, category, year, null, justExtrations);

                            }
                        }

                        excelTable.WorkSheet.Cells[cellAdress].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");

                        excelTable.WorkSheet.Cells[cellAdress].Value = categoryAverage;
                    }
                }
            }
            return excelTable;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wsSheet"></param>
        /// <param name="startRow"></param>
        /// <param name="movementsModel"></param>
        /// <param name="categories"></param>
        /// <param name="year"></param>
        /// <param name="justExtrations"></param>
        /// <returns></returns>
        public static ExcelTable CreateCategoriesMonthsAveragetest(ExcelWorksheet wsSheet, int startRow, List<TransactionViewModel> movementsModel, IEnumerable<string> categories,
            int year = 0, bool justExtrations = true)
        {
            categories = justExtrations ? ModelOperation.GetExtractionCategories(categories, movementsModel) : ModelOperation.GetIncomsCategories(categories, movementsModel);
            var newAverageColumn = new List<string>();

            var startTableCell = ExcelHelpers.AddRowAndColumnToCellAddress(_startCell, startRow, 0);

            categories = categories.OrderBy(c => c);
            var columns = new[] { "Month" };
            if (year > 0)
            {
                newAverageColumn.Add($"Year({year})");
            }
            else
            {
                newAverageColumn.Add("Years");
                var years = movementsModel.Select(mov => mov.DateTime.Year).Distinct();
                newAverageColumn.AddRange(years.Select(selectedYear => $"Year({selectedYear})"));
            }

            var newExcelColumn = columns.Concat(categories);
            // Calculate size of the table
            var excelColumns = newExcelColumn as string[] ?? newExcelColumn.ToArray();
            var endTableCell = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, 12, excelColumns.Count() - 1);

            //give table Name
            var tableName = "CategoriesMonthAverage1";
            // Add table Headers
            var excelTable = CreateExcelTable(wsSheet, tableName, excelColumns, startTableCell, endTableCell, true);
            for (int month = 1; month <= 12; month++)
            {
                for (int column = 0; column < excelColumns.Count(); column++)
                {
                    var tableCell = ExcelHelpers.AddRowAndColumnToCellAddress(startTableCell, month, column);
                    if (excelColumns[column] == "Month")
                    {
                        var monthName = string.Concat(DateTimeFormatInfo.CurrentInfo.GetMonthName(month));
                        excelTable.WorkSheet.Cells[tableCell].Value = monthName;
                    }
                    else
                    {
                        excelTable.WorkSheet.Cells[tableCell].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");

                        var average = ModelOperation.AverageforCategory(movementsModel, excelColumns[column], year, month, justExtrations);
                        excelTable.WorkSheet.Cells[tableCell].Value = average;
                    }
                }
            }
            return excelTable;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelTable"></param>
        /// <param name="columnName"></param>
        /// <param name="cellAdress"></param>
        public static void UpdateTableValues(ExcelTable excelTable, string columnName, string cellAdress)
        {
            var columnNameAdresse = ExcelHelpers.GetAdressFromColumnName(excelTable, columnName);

            for (int month = 1; month <= 12; month++)
            {
                var cellToUpdate = ExcelHelpers.AddRowAndColumnToCellAddress(columnNameAdresse, month, 0);
                var totalCellValue = ExcelHelpers.AddRowAndColumnToCellAddress(cellAdress, month, 0);
                excelTable.WorkSheet.Cells[cellToUpdate].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                excelTable.WorkSheet.Cells[cellToUpdate].Formula = totalCellValue;
            }
        }
        public static void UpdateClassesTableValues(Dictionary<string, string> BudgetCategoriesAddressdictionary, Dictionary<string, string> ExpenseCategoriesAddressdictionary, ExcelTable excelTable)
        {
            //TODO check cell Value
            var date = (double)excelTable.WorkSheet.Cells["G1"].Value;
            var monthToFilter = DateTime.FromOADate(date).Month;

            // Get cell address
            var operatingAdress = ExcelHelpers.GetAdressFromColumnName(excelTable, "OPERATING");
            var budgetAdress = ExcelHelpers.GetAdressFromColumnName(excelTable, "BUDGET");
            var actualAdress = ExcelHelpers.GetAdressFromColumnName(excelTable, "ACTUAL");

            var categories = BudgetCategoriesAddressdictionary.Where(ct => ExpenseCategoriesAddressdictionary.ContainsKey(ct.Key)).Select(ct => ct.Key).ToList();
            var tableElement = excelTable.TableXml.DocumentElement;

            //tableElement.Attributes["ref"].Value = rng.Address;
            var ref1 = tableElement.Attributes["ref"].Value;

            var columnNode = tableElement["tableColumns"];
            //columnNode.Attributes["count"].Value = rng.End.Column.ToString();
            var count1 = columnNode.Attributes["count"].Value;

            for (int row = 0; row < categories.Count; row++)
            {
                //TODO Update formula =HLOOKUP([@OPERATING];'Expenses details'!$E$22:$AC$34;MONTH($G$1)+1;FALSE)

                string budgetCellAdress = BudgetCategoriesAddressdictionary[categories[row]];
                string actualCellAdress = ExpenseCategoriesAddressdictionary[categories[row]];

                string newBudgetCell = $"OFFSET({budgetCellAdress},MONTH($G$1),0)";
                string newActualCell = $"OFFSET({actualCellAdress},MONTH($G$1),0)";

                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(operatingAdress, row + 1, 0)].Value = categories[row];

                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(budgetAdress, row + 1, 0)].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(budgetAdress, row + 1, 0)].Formula = newBudgetCell;

                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(actualAdress, row + 1, 0)].Style.Numberformat.Format = ExcelHelpers.SetFormatToCell("Amount");
                excelTable.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(actualAdress, row + 1, 0)].Formula = newActualCell;

            }

        }


        // TODO controll




        public static void AddSheetHeading(ExcelWorksheet wsSheet, string tableTitle)
        {
            wsSheet.Cells[_hedingCell].Value = "Table Name";
            wsSheet.Cells[_titlecell].Value = tableTitle;
            wsSheet.Cells[_hedingCell].Style.Font.Size = 12;
            wsSheet.Cells[_hedingCell].Style.Font.Bold = true;
            wsSheet.Cells[_hedingCell].Style.Font.Italic = true;
        }



        private static ExcelTable CreateExcelTable(ExcelWorksheet wsSheet, string tableName, IEnumerable<string> excelColumns, string startTableCellAddress, string endTableCellAddress, bool showTotal = false)
        {
            ExcelTable table;
            using (ExcelRange rng = wsSheet.Cells[$"{startTableCellAddress}:{endTableCellAddress}"])
            {
                //Indirectly access ExcelTableCollection class
                table = wsSheet.Tables.Add(rng, tableName);
            }

            //var color = Color.FromArgb(250, 199, 111);

            //Set Columns position & name
            var i = 0;
            foreach (var property in excelColumns)
            {
                table.Columns[i].Name = string.Concat(property);
                //Add Subtotal cell to the end of the table
                if (i != 0)
                    table.Columns[i].TotalsRowFormula = $"SUBTOTAL(101,[{property}])"; // 101 average, 103 Count, 109 sum
                i++;
            }

            table.ShowHeader = true;
            table.ShowFilter = true;
            table.ShowTotal = showTotal;
            return table;
        }

        //
        public static void UpdateBudgetCashFlow(ExcelPackage excelPackage, List<TransactionViewModel> movementsViewModels, List<string> categoriesArray, int year)
        {
            ExcelTable yearBudgetTable = null;
            ExcelTable yearExpensesTable = null;
            if (year == 0)
            {
                year = DateTime.Today.Year;
            }

            // Create Cashflow
            var expensesWSheet = excelPackage.Workbook.Worksheets["Expenses details"];

            // add year categoiers Table
            try
            {
                var yearExpensesTables = ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(expensesWSheet, movementsViewModels, categoriesArray, year, "YearExpenses", true, "B38");
                yearExpensesTable = yearExpensesTables.FirstOrDefault();
            }
            catch (Exception e)
            {
                throw new Exception("Creating Year expensesTable Sheet. Error message : " + e.Message);
            }



            // add year incoms categoiers 
            try
            {
                var yearIncomsTables = ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(expensesWSheet, movementsViewModels, categoriesArray, year, "YearIncoms", false, "B54");
            }
            catch (Exception e)
            {
                throw new Exception("Problem Creating Year IncomsTable Sheet. Error message : " + e.Message);
            }

            // update Year table
            Dictionary<string, string> categoriesAddressWithTotals = null;
            Dictionary<string, string> categoriesAddress = null;
            //Add sub total and total to list to get them from budget table
            try
            {
                var categoryListWithTotals = Helpers.AddItemsToIenumeration(categoriesArray, new List<string>() { "Sub Total", "Total" });
                yearBudgetTable = expensesWSheet.Tables["Year_budget"];
                var listWithTotals = categoryListWithTotals as string[] ?? categoryListWithTotals.ToArray();
                if (yearBudgetTable != null)
                {
                    categoriesAddressWithTotals = ExcelHelpers.GetNamesAdress(listWithTotals, yearBudgetTable);
                }

                //Get address to expenses table
                categoriesAddress = ExcelServices.GetColumnsNameAdress(listWithTotals, yearExpensesTable);

            }
            catch (Exception e)
            {
                throw new Exception("Cant get Info from table from 'Expenses details' sheet. Error message : " + e.Message);
            }

            //Update year excel table
            try
            {
                var yearWSheet = excelPackage.Workbook.Worksheets["Year summary"];
                var tblOperatingExpensesTable = yearWSheet.Tables["tblOperatingExpenses"];
                string keyCellValue = null;
                if (categoriesAddressWithTotals != null)
                {
                    if (categoriesAddressWithTotals.TryGetValue("Total", out keyCellValue))
                    {
                        ExcelServices.UpdateTableValues(tblOperatingExpensesTable, "BUDGET", keyCellValue);
                    }
                }

                if (categoriesAddress != null)
                {
                    if (categoriesAddress.TryGetValue("Total", out keyCellValue))
                    {
                        ExcelServices.UpdateTableValues(tblOperatingExpensesTable, "ACTUAL", keyCellValue);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Cant tables in 'Year summary' Table values. Error message : " + e.Message);
            }

            Dictionary<string, string> monthBudgetCategoriesAddress = null;
            Dictionary<string, string> monthExpensesCategoriesAddress = null;
            try
            {
                // get address to Month budget table
                var categoriesWithoutIncome = Helpers.DeleteItemsfromIenumeration(categoriesArray, new List<string>() { "Åse", "Matias" });
                monthBudgetCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearBudgetTable);
                monthExpensesCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearExpensesTable);
            }
            catch (Exception e)
            {
                throw new Exception("Cant get tables info from 'Expenses details' sheet to update Class table. Error message : " + e.Message);
            }

            //update month Table with the categories summary
            try
            {
                var monthWSheet = excelPackage.Workbook.Worksheets["Monthly summary"];
                var tblOperatingExpenses7Table = monthWSheet.Tables["tblOperatingExpenses7"];
                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, tblOperatingExpenses7Table);
            }
            catch (Exception e)
            {
                throw new Exception("Cant update tblOperatingExpenses7 in 'Monthly summary' sheet. Error message : " + e.Message);
            }

            //return excelPackage;
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
    }
}

