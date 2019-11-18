using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using FluentAssertions;
using homeBudget.Models;
using homeBudget.Services;
using OfficeOpenXml;
using Xunit;

namespace homeBudget.Tests
{
    public class ExcelServicesTests
    {
        [Fact]
        public void CreateExcelMonthSummaryTableFromMovementsViewModelTest()
        {

            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> movementsModels = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            movementsModels.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();


            var movementsViewModels = ModelConverter.CreateMovementsViewModels(movementsModels, categorisModel, "Felles");

            movementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

            ExcelWorksheet wsSheet;
            IEnumerable<string> categories;
            ExcelPackage excelPackage = new ExcelPackage();

            using (var stream = new MemoryStream())
            using (var package = new ExcelPackage(stream))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                wsSheet = package.Workbook.Worksheets["Sheet1"];

                ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(wsSheet, movementsViewModels, categoryList, 0, "Table-Test", true);

                //now test that it generated properly, such as:
                var noko = package.Workbook.Worksheets.FirstOrDefault();
                var noko2 = package.Workbook.Worksheets["Sheet1"].Cells["A3"].Value;
                var noko3 = package.Workbook.Worksheets["Sheet1"].Cells["A4"].Value;

                var saveExcel = TestsHelper.SaveExcrlPackage(package, "test-Temp2");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void CreateExcelTableFromMovementsViewModelTest()
        {

            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> movementsModels = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            movementsModels.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();


            var movementsViewModels = ModelConverter.CreateMovementsViewModels(movementsModels, categorisModel, "Felles");

            movementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

            ExcelWorksheet wsSheet;
            IEnumerable<string> categories;
            ExcelPackage excelPackage = new ExcelPackage();

            using (var stream = new MemoryStream())
            using (var package = new ExcelPackage(stream))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                wsSheet = package.Workbook.Worksheets["Sheet1"];

                ExcelServices.CreateExcelTableFromMovementsViewModel(movementsViewModels, wsSheet, "Transactions");


                //now test that it generated properly, such as:
                var sheeDefault = package.Workbook.Worksheets.FirstOrDefault();
                if (sheeDefault != null) sheeDefault.Cells["D57"].Value.Should().Be(-55);

                var saveExcel = TestsHelper.SaveExcrlPackage(package, "test-Temp");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void CreateAverageForYearMonthDayTest()
        {

            IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
            var movementsViewModels = TestsHelper.GetMovementsViewModels();

            using (var stream = new MemoryStream())
            using (var package = new ExcelPackage(stream))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                var categoriesAverageWSheet = package.Workbook.Worksheets["Sheet1"];

                ExcelServices.CreateAverageForYearMonthDay(movementsViewModels, categoriesAverageWSheet, categoryList, 0, 0, true);

                //now test that it generated properly, such as:
                //var sheeDefault = package.Workbook.Worksheets.FirstOrDefault();
                //if (sheeDefault != null) sheeDefault.Cells["D57"].Value.Should().Be(-55);

                var saveExcel = TestsHelper.SaveExcrlPackage(package, "CreateAverageForYearMonthDay-test");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void CreateCategoriesMonthsAverageTest()
        {
            IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
            var movementsViewModels = TestsHelper.GetMovementsViewModels();

            using (var stream = new MemoryStream())
            using (var package = new ExcelPackage(stream))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                var categoriesAverageWSheet = package.Workbook.Worksheets["Sheet1"];

                var excelTable = ExcelServices.CreateCategoriesMonthsAveragetest(categoriesAverageWSheet, 13, movementsViewModels, categoryList, 2019, true);
                var table = package.Workbook.Worksheets.FirstOrDefault().Tables.FirstOrDefault();

                var start = table.Address.Start.Address;
                var end = table.Address.End.Address;

                var noko = table.WorkSheet.Cells["B6"].Value = "";
                //now test that it generated properly, such as:
                //var sheeDefault = package.Workbook.Worksheets.FirstOrDefault();
                //if (sheeDefault != null) sheeDefault.Cells["D57"].Value.Should().Be(-55);

                var saveExcel = TestsHelper.SaveExcrlPackage(package, "CreateCategoriesMonthsAveragetest-test");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void CreateExcelMonthSummaryTableFromMovementsViewModelIncomeTest()
        {
            IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
            var movementsViewModels = TestsHelper.GetMovementsViewModels();

            using (var stream = new MemoryStream())
            using (var package = new ExcelPackage(stream))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                var categoriesAverageWSheet = package.Workbook.Worksheets["Sheet1"];

                var excelTable = ExcelServices.CreateCategoriesMonthsAveragetest(categoriesAverageWSheet, 13, movementsViewModels, categoryList, 2019, false);
                var table = package.Workbook.Worksheets.FirstOrDefault().Tables.FirstOrDefault();

                var start = table.Address.Start.Address;
                var end = table.Address.End.Address;

                var noko = table.WorkSheet.Cells["B6"].Value = "";
                //now test that it generated properly, such as:
                //var sheeDefault = package.Workbook.Worksheets.FirstOrDefault();
                //if (sheeDefault != null) sheeDefault.Cells["D57"].Value.Should().Be(-55);

                var saveExcel = TestsHelper.SaveExcrlPackage(package, "Incomes-test");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void UpdateTableValuesBudgetTest()
        {
            Dictionary<string, string> CategoriesAddressWithTotals = null;

            var streamFile = TestsHelper.GetAssemblyFile("Budget Cashflow.xlsx");
            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
                var categoryListWithTotals = Helpers.AddItemsToIenumeration(categoryList, new List<string>() { "Sub Total", "Total" });
                var ExpensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];


                var yearBudgetTable = ExpensesWSheet.Tables["Year_budget"];
                if (yearBudgetTable != null)
                {
                    CategoriesAddressWithTotals = ExcelHelpers.GetNamesAdress(categoryListWithTotals, yearBudgetTable);

                }

                var yearWSheet = cashflowExcelPkg.Workbook.Worksheets["Year summary"];
                var excelTable = yearWSheet.Tables["tblOperatingExpenses"];
                string keyCellValue = null;
                if (CategoriesAddressWithTotals != null)
                {

                    if (CategoriesAddressWithTotals.TryGetValue("Total", out keyCellValue))
                    {
                        ExcelServices.UpdateTableValues(excelTable, "BUDGET", keyCellValue);
                    }
                }
                var saveExcel = TestsHelper.SaveExcrlPackage(cashflowExcelPkg, "Update-Test1");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void UpdateTableValuesActualTest()
        {
            Dictionary<string, string> CategoriesAddress = null;

            var streamFile = TestsHelper.GetAssemblyFile("Budget Cashflow.xlsx");
            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
                var categoryListWithTotals = Helpers.AddItemsToIenumeration(categoryList, new List<string>() { "Sub Total", "Total" });
                var ExpensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];

                var yearBudgetTable = ExpensesWSheet.Tables["Table_2018"];
                if (yearBudgetTable != null)
                {
                    CategoriesAddress = ExcelHelpers.GetNamesAdress(categoryListWithTotals, yearBudgetTable);

                }

                var yearWSheet = cashflowExcelPkg.Workbook.Worksheets["Year summary"];
                var excelTable = yearWSheet.Tables["tblOperatingExpenses"];
                string keyCellValue = null;
                if (CategoriesAddress != null)
                {
                    if (CategoriesAddress.TryGetValue("Total", out keyCellValue))
                    {
                        ExcelServices.UpdateTableValues(excelTable, "ACTUAL", keyCellValue);
                    }
                }
                var saveExcel = TestsHelper.SaveExcrlPackage(cashflowExcelPkg, "Update-Test2");
                saveExcel.Should().BeTrue();
            }
        }
        [Fact]
        public void UpdateClassesTableValues()
        {
            var streamFile = TestsHelper.GetAssemblyFile("Budget Cashflow.xlsx");
            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
                var expensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];

                var yearBudgetTable = expensesWSheet.Tables["Year_budget"];
                var yearExpensesTable = expensesWSheet.Tables["Table_2018"];

                var monthWSheet = cashflowExcelPkg.Workbook.Worksheets["Monthly summary"];
                var tblOperatingExpenses7Table = monthWSheet.Tables["tblOperatingExpenses7"];


                Dictionary<string, string> monthBudgetCategoriesAddress = null;
                Dictionary<string, string> monthExpensesCategoriesAddress = null;


                // get address to Month budget table
                var categoriesWithoutIncome = Helpers.DeleteItemsfromIenumeration(categoryList, new List<string>() { "Åse", "Matias" });
                monthBudgetCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearBudgetTable);
                monthExpensesCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearExpensesTable);
                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, tblOperatingExpenses7Table);

                var saveExcel = TestsHelper.SaveExcrlPackage(cashflowExcelPkg, "UpdateClassesTableValues-Test");
                saveExcel.Should().BeTrue();
            }
        }

        //Comun to all

    }
}
