using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Linq;
using System.Reflection;
using homeBudget.Models;
using homeBudget.Services;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml.Table;
using Newtonsoft.Json;

namespace homeBudget.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    public class TransactionsController : Controller
    {
        private IHostingEnvironment _hostingEnvironment;

        public TransactionsController(IHostingEnvironment environment)
        {
            _hostingEnvironment = environment;
        }

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Post(IFormFile transactions, IFormFile categories, int year = 0)
        {

            var filePathTemp = Path.GetTempFileName();
            var filePath1 = Path.GetTempFileName();
            var responseDictionary = new Dictionary<string, string>();

            if (!IsFileValid(transactions) || !IsFileValid(categories))
                return BadRequest();

            ExcelWorksheet transactionsWorkSheet;
            ExcelWorksheet categoriesWorkSheet;

            //Read Excel Files
            try
            {
                transactionsWorkSheet = await ExcelHelpers.GetExcelWorkSheet(transactions, filePathTemp);
                categoriesWorkSheet = await ExcelHelpers.GetExcelWorkSheet(categories, filePath1);
            }
            catch (Exception ex)
            {

                return BadRequest("Can't read excel files");
            }


            var transactionsTable = transactionsWorkSheet.Tables.FirstOrDefault();
            var categoriestabTable = categoriesWorkSheet.Tables.FirstOrDefault();
            //Get excel data  in Json format easier to serialize to class
            var accountMovmentjArray = ExcelConverter.GetJsonFromTable(transactionsTable);
            var subCategoriesjArray = ExcelConverter.GetJsonFromTable(categoriestabTable);

            // serialize Json to Class
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(accountMovmentjArray);
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();

            //TODO Get acount Name from Excel or Input variable
            var movementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            //Create excel Sheet with the transaction updated with the keewords, categories, and subproject (is exists)
            var categoriesArray = categoryList as string[] ?? categoryList.ToArray();
            using (var stream = new MemoryStream())
            using (var transactionUpdatePackage = new ExcelPackage(stream))
            {
                try
                {
                    //Add Table Title
                    var transactionSheet = transactionUpdatePackage.Workbook.Worksheets.Add("Transactions");
                    ExcelServices.AddSheetHeading(transactionSheet, "Transactions and Categories");

                    //Add transactions to excel Sheet
                    var movementsViewExcelTable = ExcelServices.CreateExcelTableFromMovementsViewModel(movementsViewModels, transactionSheet, "Transactions");

                }
                catch (Exception e)
                {
                    return BadRequest("Error Creating transaction sheet. Error message : " + e.Message);
                }

                // Add Categories Average to excel
                try
                {
                    AddCategoriesAverage(year, transactionUpdatePackage, movementsViewModels, categoriesArray);
                }
                catch (Exception e)
                {

                    return BadRequest(" Error Creating Average sheet. Error message : " + e.Message);
                }

                // add month summaries to excel
                try
                {
                    var monthSummariesSheet = transactionUpdatePackage.Workbook.Worksheets.Add("MonthSummaries");
                    ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(monthSummariesSheet, movementsViewModels, categoriesArray, 0, null, true);
                }
                catch (Exception e)
                {
                    return BadRequest("Creating MonthSummary Sheet. Error message : " + e.Message);
                }

                try
                {
                    var filename = "Transactions Update With Categories";

                    string contentRootPath = _hostingEnvironment.ContentRootPath;

                    var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                    transactionUpdatePackage.SaveAs(new FileInfo(fullPath));
                }
                catch (Exception e)
                {
                    return BadRequest("Transactions Update With Categories Can't be saved :" + e.Message);
                }

            }

            //Next Excel File More details and cashflow + Chars
            using (var cashflowExcelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx")))
            {


                try
                {
                    UpdateBudgetCashFlow(cashflowExcelPkg, movementsViewModels, categoriesArray.ToList(), year);
                }
                catch (Exception)
                {
                    return BadRequest("Cant creat Cashflow Excel File");
                }
                // Save Excel Package
                try
                {
                    var filename = $"Budget Cashflow ({year})";

                    string contentRootPath = _hostingEnvironment.ContentRootPath;

                    var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                    cashflowExcelPkg.SaveAs(new FileInfo(fullPath));
                }
                catch
                {
                    return BadRequest("Can't be saved");
                }
            }

            return Ok(responseDictionary);
        }

        [HttpPost("UploadTransactions")]
        public async Task<IActionResult> PostTransaction(IFormFile transactions, int year = 0)
        {
            var filePathTemp = Path.GetTempFileName();

            if (!IsFileValid(transactions))
                return BadRequest("Can't read transactions excel file");

            using (var stream = new FileStream(filePathTemp, FileMode.Create))
            {
                await transactions.CopyToAsync(stream);
                List<TransactionViewModel> movementsViewModels = null;
                List<string> categoryList;
                using (var transacationAndCategories = new ExcelPackage(stream))
                {
                    JArray jsonFromTable = null;
                    var worksheet = transacationAndCategories.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null)
                    {
                        var excelTable = worksheet.Tables.FirstOrDefault();
                        jsonFromTable = ExcelConverter.GetJsonFromTable(excelTable);
                    }

                    if (jsonFromTable != null)
                    {
                        movementsViewModels = jsonFromTable.ToObject<List<TransactionViewModel>>();
                        //movementsViewModels = JsonConvert.DeserializeObject<List<TransactionViewModel>>(jsonFromTable?.ToString(), JsonServices.GetJsonSerializerSettings());
                    }

                    categoryList = ModelConverter.GetListOfCategories(movementsViewModels);


                    // Add Categories Average to excel
                    try
                    {
                        AddCategoriesAverage(year, transacationAndCategories, movementsViewModels, categoryList.ToArray());
                    }
                    catch (Exception e)
                    {

                        return BadRequest(" Error Creating Average sheet. Error message : " + e.Message);
                    }

                    // add month summaries to excel

                    try
                    {
                        ExcelWorksheet monthSummariesSheet = transacationAndCategories.Workbook.Worksheets["MonthSummaries"];
                        

                        if (monthSummariesSheet != null)
                        {
                            transacationAndCategories.Workbook.Worksheets.Delete("MonthSummaries");
                            monthSummariesSheet = transacationAndCategories.Workbook.Worksheets.Add("Month Summaries New");
                        }
                        else
                        {
                            monthSummariesSheet = transacationAndCategories.Workbook.Worksheets.Add("MonthSummaries");
                        }

                        ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(monthSummariesSheet, movementsViewModels, categoryList.ToArray(), 0, null, true);
                    }
                    catch (Exception e)
                    {
                        return BadRequest("Creating MonthSummary Sheet. Error message : " + e.Message);
                    }

                    try
                    {
                        var filename = "Transactions Update With Categories (1)";

                        string contentRootPath = _hostingEnvironment.ContentRootPath;

                        var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                        transacationAndCategories.SaveAs(new FileInfo(fullPath));
                    }
                    catch (Exception e)
                    {
                        return BadRequest("Transactions Update With Categories Can't be saved :" + e.Message);
                    }



                }

                using (var cashflowExcelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx")))
                {
                    try
                    {
                        UpdateBudgetCashFlow(cashflowExcelPkg, movementsViewModels, categoryList, year);
                    }
                    catch (Exception)
                    {
                        return BadRequest("Cant creat Cashflow Excel File");
                    }
                    // Save Excel Package
                    try
                    {
                        var filename = "Budget Cashflow (1)";
                        string contentRootPath = _hostingEnvironment.ContentRootPath;

                        var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                        cashflowExcelPkg.SaveAs(new FileInfo(fullPath));
                        return Ok("File " + filename + "cretated in :" + fullPath);
                    }
                    catch (Exception e)
                    {
                        return BadRequest("Can't be saved :" + e.Message);
                    }
                }
            }
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="year"></param>
        /// <param name="transactionUpdatePackage"></param>
        /// <param name="movementsViewModels"></param>
        /// <param name="categoriesArray"></param>
        private static void AddCategoriesAverage(int year, ExcelPackage transactionUpdatePackage, List<TransactionViewModel> movementsViewModels, string[] categoriesArray)
        {
            ExcelWorksheet categoriesAverageWSheet = transactionUpdatePackage.Workbook.Worksheets["Categories Average"];
            if (categoriesAverageWSheet != null)
            {
                //ExcelHelpers.DeldeteExcelTablesFromWorkSheet(categoriesAverageWSheet);
                transactionUpdatePackage.Workbook.Worksheets.Delete("Categories Average");
                categoriesAverageWSheet = transactionUpdatePackage.Workbook.Worksheets.Add("Categories Average new");
            }
            else
            {
                categoriesAverageWSheet = transactionUpdatePackage.Workbook.Worksheets.Add("Categories Average");
            }

            ExcelServices.AddSheetHeading(categoriesAverageWSheet, "Transactions and Categories");

            var yearMonthTable = ExcelServices.CreateAverageForYearMonthDay(movementsViewModels, categoriesAverageWSheet, categoriesArray, year, 0, true);

            var endTableRow = yearMonthTable.Address.End.Row;
            var categoryMonthAvgTable = ExcelServices.CreateCategoriesMonthsAveragetest(categoriesAverageWSheet, endTableRow, movementsViewModels, categoriesArray, year, true);
        }


        private void UpdateBudgetCashFlow(ExcelPackage excelPackage, List<TransactionViewModel> movementsViewModels, List<string> categoriesArray, int year)
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
    }
}