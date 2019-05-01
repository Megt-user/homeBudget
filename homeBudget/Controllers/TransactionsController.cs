using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExcelClient;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Transactions.Models;
using Transactions.Services;
using System.Linq;
using System.Reflection;
using OfficeOpenXml.Table;
using Newtonsoft.Json;

namespace homeBudget.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    public class TransactionsController : Controller
    {

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Post(IFormFile transacation, IFormFile category, int year = 0)
        {

            long categoryFileSize = category.Length;

            var filePathTemp = Path.GetTempFileName();
            var filePath1 = Path.GetTempFileName();


            if (!IsFileValid(transacation) || !IsFileValid(category))
                return BadRequest();

            ExcelWorksheet transactionsWorkSheet;
            ExcelWorksheet categoriesWorkSheet;
            
            //Start Columns positions
            new ExcelServices();

            //Read Excel Files
            using (var stream = new FileStream(filePathTemp, FileMode.Create))
            {
                await transacation.CopyToAsync(stream);
                transactionsWorkSheet = ExcelServices.GetExcelWorksheet(stream);
            }
            using (var stream = new FileStream(filePath1, FileMode.Create))
            {
                await category.CopyToAsync(stream);
                categoriesWorkSheet = ExcelServices.GetExcelWorksheet(stream);
            }

            //Get excel data  in Json format easier to serialize to class
            var subCategoriesjArray = JArray.Parse(ExcelServices.GetJsonFromTable(categoriesWorkSheet));
            var accountMovmentjArray = JArray.Parse(ExcelServices.GetJsonFromTable(transactionsWorkSheet));
            
            // serialize Json to Class
            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);
            List<SubCategory> categorisModel = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();

            //TODO Get acount Name from Excel or Input variable
            var modementsViewModels = ModelClassServices.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            var excelPkg = new ExcelPackage();
            try
            {
                ExcelServices.CreateSheetWithTransactionMovments(modementsViewModels, excelPkg, "Transactions", "Transactions and Categories", "Transactions");
            }
            catch (Exception e)
            {
                return BadRequest("Creating transaction sheet. Error message : " + e.Message);
            }
            try
            {
                ExcelServices.CreateSheetWithMonthSummary(modementsViewModels, excelPkg, "MonthSummaries", categoryList);
            }
            catch (Exception e)
            {
                return BadRequest("Creating MonthSummary Sheet. Error message : " + e.Message);
            }

            // Create Cashflow
            var cashflowExcelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx"));
            var ExpensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];

            if (year == 0)
            {
                year = DateTime.Today.Year;
            }

            // add year categoiers Table
            try
            {
                ExcelServices.CreateYearExpensesTable(modementsViewModels, categoryList, year, ExpensesWSheet, "YearExpenses", "B38");

            }
            catch (Exception e)
            {

                return BadRequest("Creating Year expensesTable Sheet. Error message : " + e.Message);

            }
            // add year incoms categoiers 
            try
            {
                ExcelServices.CreateYearIncomsTable(modementsViewModels, categoryList, year, ExpensesWSheet, "YearIncoms", "B54");

            }
            catch (Exception e)
            {
                return BadRequest("Problem Creating Year IncomsTable Sheet. Error message : " + e.Message);
            }

            // update Year table

            Dictionary<string, string> CategoriesAddressWithTotals = null;
            Dictionary<string, string> CategoriesAddress = null;
            IEnumerable<string> categoryListWithTotals = null;
            //Add sub total and total to list to get them from budget table
            try
            {
                categoryListWithTotals = Helpers.AddItemsToIenumeration(categoryList, new List<string>() { "Sub Total", "Total" });
                CategoriesAddressWithTotals = ExcelServices.GetColumnsNameAdress(categoryListWithTotals, ExpensesWSheet, "Year_budget");
                //Get address to expenses table
                CategoriesAddress = ExcelServices.GetColumnsNameAdress(categoryListWithTotals, ExpensesWSheet, "YearExpenses");

            }
            catch (Exception e)
            {

                return BadRequest("Cant get Info from table from 'Expenses details' sheet. Error message : " + e.Message);

            }

            //Update year excel table
            try
            {
                var yearWSheet = cashflowExcelPkg.Workbook.Worksheets["Year summary"];

                ExcelServices.UpdateYearTableValues(CategoriesAddressWithTotals, year, yearWSheet, "tblOperatingExpenses", "BUDGET", "Total");
                ExcelServices.UpdateYearTableValues(CategoriesAddress, year, yearWSheet, "tblOperatingExpenses", "ACTUAL", "Total");

            }
            catch (Exception e)
            {

                return BadRequest("Cant tables in 'Year summary' sheet. Error message : " + e.Message);

            }


            Dictionary<string, string> monthBudgetCategoriesAddress = null;
            Dictionary<string, string> monthExpensesCategoriesAddress = null;
            try
            {
                // get address to Month budget table
                var categoriesWithoutIncome = Helpers.DeleteItemsfromIenumeration(categoryList, new List<string>() { "Åse", "Matias" });
                monthBudgetCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, ExpensesWSheet, "Year_budget");
                monthExpensesCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, ExpensesWSheet, "YearExpenses");

            }
            catch (Exception e)
            {

                return BadRequest("Cant get tables info from 'Expenses details' sheet to update Class table. Error message : " + e.Message);

            }
            //update month Table with the categories summary
            try
            {
                var monthWSheet = cashflowExcelPkg.Workbook.Worksheets["Monthly summary"];
                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, year, monthWSheet, "tblOperatingExpenses7");

            }
            catch (Exception e)
            {
                return BadRequest("Cant update tblOperatingExpenses7 in 'Monthly summary' sheet. Error message : " + e.Message);
            }

            Dictionary<string, string> filesPath = new Dictionary<string, string>();
            // Save Excel Package
            try
            {
                var filename = "Transactions Update With Categories";
                var path = Path.Combine(@"h:\", "Transactions");

                Directory.CreateDirectory(path);
                excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(filename, ".xlsx"))));
                filesPath.Add(filename, path);
                filename = $"Budget Cashflow ({year})";
                path = Path.Combine(@"h:\", "Transactions");
                excelPkg.Dispose();

                Directory.CreateDirectory(path);
                cashflowExcelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(filename, ".xlsx"))));
                filesPath.Add(filename, path);
                cashflowExcelPkg.Dispose();
            }
            catch
            {

                return BadRequest("Can't be saved");
            }

            return Ok("Created in:" + filesPath);
        }

        [HttpPost("UploadTransactions")]
        public async Task<IActionResult> PostTransaction(IFormFile transacation, int year = 0)
        {


            long categoryFileSize = transacation.Length;

            var filePathTemp = Path.GetTempFileName();
            var filePath1 = Path.GetTempFileName();


            if (!IsFileValid(transacation))
                return BadRequest("Can't be saved");
            ExcelWorksheet transactionsWorkSheet;

            using (var stream = new FileStream(filePathTemp, FileMode.Create))
            {
                await transacation.CopyToAsync(stream);
                transactionsWorkSheet = ExcelServices.GetExcelWorksheet(stream);
            }
            var WoorSheet = transactionsWorkSheet.Workbook.Worksheets.FirstOrDefault();
            var jsonFromTable = ExcelServices.GetJsonFromTable(WoorSheet);
            transactionsWorkSheet.Dispose();

            List<MovementsViewModel> movementsViewModels = JsonConvert.DeserializeObject<List<MovementsViewModel>>(jsonFromTable, JsonServices.GetJsonSerializerSettings());

            var categoryList = ModelClassServices.GetListOfCategories(movementsViewModels);

            //var movements = JsonConvert.DeserializeObject<List<MovementsViewModel>>(jsonFromTable, dateTimeConverter);
            var excelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx"));
            try
            {
                var ExpensesWSheet = excelPkg.Workbook.Worksheets["Expenses details"];

                if (year == 0)
                {
                    year = DateTime.Today.Year;
                }

                //workSheet.Tables.Delete("YearExpenses");

                // add all year categoiers 
                ExcelServices.CreateYearExpensesTable(movementsViewModels, categoryList, year, ExpensesWSheet, "YearExpenses", "B38");

                // update Year table

                //Get Adress to budget table
                var categoryListWithTotals = Helpers.AddItemsToIenumeration(categoryList, new List<string>() { "Sub Total", "Total" });
                var CategoriesAddressWithTotals = ExcelServices.GetColumnsNameAdress(categoryListWithTotals, ExpensesWSheet, "Year_budget");

                //Get address to expenses table
                var CategoriesAddress = ExcelServices.GetColumnsNameAdress(categoryListWithTotals, ExpensesWSheet, "YearExpenses");

                //Update year excel table
                var yearWSheet = excelPkg.Workbook.Worksheets["Year summary"];

                ExcelServices.UpdateYearTableValues(CategoriesAddressWithTotals, year, yearWSheet, "tblOperatingExpenses", "BUDGET", "Total");
                ExcelServices.UpdateYearTableValues(CategoriesAddress, year, yearWSheet, "tblOperatingExpenses", "ACTUAL", "Total");

                // get address to Month budget table
                var categoriesWithoutIncome = Helpers.DeleteItemsfromIenumeration(categoryList, new List<string>() { "Åse", "Matias" });
                var monthBudgetCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, ExpensesWSheet, "Year_budget");
                var monthExpensesCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, ExpensesWSheet, "YearExpenses");

                //update month Table with the categories summary
                var monthWSheet = excelPkg.Workbook.Worksheets["Monthly summary"];

                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, year, monthWSheet, "tblOperatingExpenses7");

            }
            catch (Exception e)
            {
                var noko = e.Message;
            }
            Dictionary<string, string> filesPath = new Dictionary<string, string>();
            // Save Excel Package
            try
            {
                var filename = Path.GetFileNameWithoutExtension(transacation.FileName);
                var path = Path.Combine(@"h:\", "Transactions");

                Directory.CreateDirectory(path);
                excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat($"{filename}_New.xlsx"))));
                filesPath.Add(filename, path);
            }
            catch
            {

                return BadRequest("Can't be saved");
            }

            return Ok("Created in:" + filesPath);

        }

        private static Stream GetAssemblyFile(string fileName)
        {
            var resourceFileNema = fileName;
            //var resourcePath = $"ExcelClient.Tests.TestData.{resourceFileNema}";
            var resourcePath = $"homeBudget.Data.{resourceFileNema}";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);
            return resourceAsStream;
        }

        private bool IsFileValid(IFormFile file)
        {
            long size = file.Length;
            string extension = Path.GetExtension(file.FileName).ToLower();
            if (size > 0)
            {
                if (extension == ".xls" || extension == ".xlsx")
                {
                    return true;
                }
            }
            return false;
        }
    }
}