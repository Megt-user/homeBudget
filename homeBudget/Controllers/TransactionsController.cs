using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using homeBudget.Services.Logger;
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
        private ILogEntryService _logEntryService;
        public TransactionsController(IHostingEnvironment environment, ILogEntryService logEntryService)
        {
            _hostingEnvironment = environment;
            _logEntryService = logEntryService;
        }

        public TransactionsController(IHostingEnvironment environment)
        {
            _hostingEnvironment = environment;
        }

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Post(IFormFile transactions, IFormFile categories, int year = 0)
        {
            var stopWatch = new Stopwatch();
            var filePathTemp = Path.GetTempFileName();
            var filePath1 = Path.GetTempFileName();
            var responseDictionary = new Dictionary<string, string>();

            stopWatch.Start();
            if (!IsFileValid(transactions) || !IsFileValid(categories))
                return BadRequest();

            ExcelWorksheet transactionsWorkSheet;
            ExcelWorksheet categoriesWorkSheet;
            LogEntry logEntry;
            //Read Excel Files
            try
            {
                transactionsWorkSheet = await ExcelHelpers.GetExcelWorkSheet(transactions, filePathTemp);
                categoriesWorkSheet = await ExcelHelpers.GetExcelWorkSheet(categories, filePath1);
                logEntry = new LogEntry("Read Excel Files", "GetExcelWorkSheet", stopWatch.ElapsedMilliseconds, "Info");


            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"error:{ex.Message}", "GetExcelWorkSheet", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                return BadRequest("Can't read excel files");
            }

            _logEntryService.Save(logEntry);
            stopWatch = Stopwatch.StartNew();
            var totalStopWatch = new Stopwatch();
            totalStopWatch.Start();

            var transactionsTable = transactionsWorkSheet.Tables.FirstOrDefault();
            var categoriestabTable = categoriesWorkSheet.Tables.FirstOrDefault();


            //Get excel data  in Json format easier to serialize to class
            var accountMovmentjArray = ExcelConverter.GetJsonFromTable(transactionsTable);
            logEntry = new LogEntry("Get Transaction Json Array", "GetJsonFromTable", stopWatch.ElapsedMilliseconds, "Info");
            _logEntryService.Save(logEntry);
            stopWatch = Stopwatch.StartNew();
            var subCategoriesjArray = ExcelConverter.GetJsonFromTable(categoriestabTable);
            logEntry = new LogEntry("Get CategoriestabTable Json Array", "GetJsonFromTable", stopWatch.ElapsedMilliseconds, "Info");
            _logEntryService.Save(logEntry);


            // serialize Json to Class
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(accountMovmentjArray);
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();

            //TODO Get acount Name from Excel or Input variable
            var movementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            //Create excel Sheet with the transaction updated with the keewords, categories, and subproject (is exists)
            var categoriesArray = categoryList as string[] ?? categoryList.ToArray();
            stopWatch = Stopwatch.StartNew();
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
                    logEntry = new LogEntry("Add transactions to excel Sheet", "CreateExcelTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);

                }
                catch (Exception ex)
                {
                    logEntry = new LogEntry($"error :{ex.Message}", "CreateExcelTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry);
                    return BadRequest("Error Creating transaction sheet. Error message");
                }

                stopWatch = Stopwatch.StartNew();
                // Add Categories Average to excel
                try
                {
                    AddCategoriesAverage(year, transactionUpdatePackage, movementsViewModels, categoriesArray);
                    logEntry = new LogEntry("Add Categories average to excel Sheet", "AddCategoriesAverage", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);
                }
                catch (Exception ex)
                {
                    logEntry = new LogEntry($"error :{ex.Message}", "AddCategoriesAverage", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry);
                    return BadRequest(" Error Creating Average sheet. Error message");
                }

                stopWatch = Stopwatch.StartNew();
                // add month summaries to excel
                try
                {
                    var monthSummariesSheet = transactionUpdatePackage.Workbook.Worksheets.Add("MonthSummaries");
                    ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(monthSummariesSheet, movementsViewModels, categoriesArray, 0, null, true);
                    logEntry = new LogEntry("Add month summaries to excel Sheet", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);
                }
                catch (Exception ex)
                {
                    logEntry = new LogEntry($"error :{ex.Message}", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry);
                    return BadRequest("Creating MonthSummary Sheet. Error message");
                }

                stopWatch = Stopwatch.StartNew();
                try
                {
                    var filename = "Transactions Update With Categories";

                    string contentRootPath = _hostingEnvironment.ContentRootPath;

                    var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                    transactionUpdatePackage.SaveAs(new FileInfo(fullPath));
                    logEntry = new LogEntry("Save excel Package", "SaveAs", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);

                    logEntry = new LogEntry("Total time procesing 'Transactions Update With Categories' Excel package", "transactionUpdatePackage", totalStopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);

                }
                catch (Exception ex)
                {
                    logEntry = new LogEntry($"error :{ex.Message}", "CreateExcelTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry); 
                    return BadRequest("Transactions Update With Categories Can't be saved");
                }

            }
            stopWatch = Stopwatch.StartNew();

            //Next Excel File More details and cashflow + Chars
            using (var cashflowExcelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx")))
            {
                logEntry = new LogEntry("Read 'Budget Cashflow.xlsx' ", "GetAssemblyFile", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);

                totalStopWatch = Stopwatch.StartNew();
                stopWatch = Stopwatch.StartNew();

                try
                {
                    UpdateBudgetCashFlow(cashflowExcelPkg, movementsViewModels, categoriesArray.ToList(), year);
                    logEntry = new LogEntry("Update budget cash flow excel Sheet", "UpdateBudgetCashFlow", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);
                }
                catch (Exception ex)
                {
                    logEntry = new LogEntry($"error :{ex.Message}", "UpdateBudgetCashFlow", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry);
                    return BadRequest("Cant creat Cashflow Excel File");
                }
                stopWatch = Stopwatch.StartNew();
                // Save Excel Package
                try
                {
                    var filename = $"Budget Cashflow ({year})";

                    string contentRootPath = _hostingEnvironment.ContentRootPath;

                    var fullPath = Path.Combine(contentRootPath, "DataTemp", $"{filename}.xlsx");
                    cashflowExcelPkg.SaveAs(new FileInfo(fullPath));
                    logEntry = new LogEntry("Save 'Budget Cashflow' excel Package", "SaveAs", stopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);

                    logEntry = new LogEntry("Total time procesing 'Budget Cashflow' Excel package", "transactionUpdatePackage", totalStopWatch.ElapsedMilliseconds, "Info");
                    _logEntryService.Save(logEntry);
                }
                catch(Exception ex)
                {
                    logEntry = new LogEntry($"Budget Cashflow - error :{ex.Message}", "CreateExcelTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                    _logEntryService.Save(logEntry);
                    return BadRequest("Budget Cashflow Can't be saved");
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
                List<MovementsViewModel> movementsViewModels = null;
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
                        movementsViewModels = jsonFromTable.ToObject<List<MovementsViewModel>>();
                        //movementsViewModels = JsonConvert.DeserializeObject<List<MovementsViewModel>>(jsonFromTable?.ToString(), JsonServices.GetJsonSerializerSettings());
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
                var monthWSheet = excelPackage.Workbook.Worksheets["Monthly summary"];
                var tblOperatingExpenses7Table = monthWSheet.Tables["tblOperatingExpenses7"];
                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, tblOperatingExpenses7Table);
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
        private static void AddCategoriesAverage(int year, ExcelPackage transactionUpdatePackage, List<MovementsViewModel> movementsViewModels, string[] categoriesArray)
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


        private void UpdateBudgetCashFlow(ExcelPackage excelPackage, List<MovementsViewModel> movementsViewModels, List<string> categoriesArray, int year)
        {
            ExcelTable yearBudgetTable = null;
            ExcelTable yearExpensesTable = null;
            if (year == 0)
            {
                year = DateTime.Today.Year;
            }

            var stopWatch = Stopwatch.StartNew();

            // Create Cashflow
            var expensesWSheet = excelPackage.Workbook.Worksheets["Expenses details"];
            LogEntry logEntry;
            // add year categoiers Table
            try
            {
                var yearExpensesTables = ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(expensesWSheet, movementsViewModels, categoriesArray, year, "YearExpenses", true, "B38");
                yearExpensesTable = yearExpensesTables.FirstOrDefault();
                logEntry = new LogEntry("Create 'YearExpenses' Table/s by MonthSummary from TransactionsViewModel", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"YearExpenses - error :{ex.Message}", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                throw new Exception("Creating Year expensesTable Sheet. Error message");
            }

            stopWatch = Stopwatch.StartNew();
            // add year incoms categoiers 
            try
            {
                var yearIncomsTables = ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(expensesWSheet, movementsViewModels, categoriesArray, year, "YearIncoms", false, "B54");
                logEntry = new LogEntry("Create 'YearIncoms' Table/s by MonthSummary from TransactionsViewModel", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Info");

                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"YearIncoms - error :{ex.Message}", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                throw new Exception("Problem Creating Year IncomsTable Sheet. Error message");
            }

            // update Year table
            Dictionary<string, string> categoriesAddressWithTotals = null;
            Dictionary<string, string> categoriesAddress = null;
            stopWatch = Stopwatch.StartNew();

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

                logEntry = new LogEntry("Get table Year_budget categories adress", "GetNamesAdress", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"Year_budget - error :{ex.Message}", "GetNamesAdress", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                throw new Exception("Cant get Info from table from 'Expenses details' sheet. Error message");
            }

            stopWatch = Stopwatch.StartNew();
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
                logEntry = new LogEntry("Update tblOperatingExpenses excel table.", "UpdateTableValues", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"tblOperatingExpenses excel table - error :{ex.Message}", "CreateExcelMonthSummaryTableFromMovementsViewModel", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                throw new Exception("Cant tables in 'Year summary' Table values. Error message");
            }

            stopWatch = Stopwatch.StartNew();
            Dictionary<string, string> monthBudgetCategoriesAddress = null;
            Dictionary<string, string> monthExpensesCategoriesAddress = null;
            try
            {
                // get address to Month budget table
                var categoriesWithoutIncome = Helpers.DeleteItemsfromIenumeration(categoriesArray, new List<string>() { "Åse", "Matias" });
                monthBudgetCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearBudgetTable);
                monthExpensesCategoriesAddress = ExcelServices.GetColumnsNameAdress(categoriesWithoutIncome, yearExpensesTable);
                logEntry = new LogEntry("Get address to Month budget table", "GetColumnsNameAdress", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"YearExpenses - error :{ex.Message}", "GetColumnsNameAdress", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry); 
                throw new Exception("Cant get tables info from 'Expenses details' sheet to update Class table. Error message");
            }
            stopWatch = Stopwatch.StartNew();
            //update month Table with the categories summary
            try
            {
                var monthWSheet = excelPackage.Workbook.Worksheets["Monthly summary"];
                var tblOperatingExpenses7Table = monthWSheet.Tables["tblOperatingExpenses7"];
                ExcelServices.UpdateClassesTableValues(monthBudgetCategoriesAddress, monthExpensesCategoriesAddress, tblOperatingExpenses7Table);
                logEntry = new LogEntry("Update tblOperatingExpenses7 Table", "UpdateClassesTableValues", stopWatch.ElapsedMilliseconds, "Info");
                _logEntryService.Save(logEntry);
            }
            catch (Exception ex)
            {
                logEntry = new LogEntry($"tblOperatingExpenses7 - error :{ex.Message}", "UpdateClassesTableValues", stopWatch.ElapsedMilliseconds, "Error");
                _logEntryService.Save(logEntry);
                throw new Exception("Cant update tblOperatingExpenses7 in 'Monthly summary' sheet. Error message");
            }

            //return excelPackage;
        }
    }
}