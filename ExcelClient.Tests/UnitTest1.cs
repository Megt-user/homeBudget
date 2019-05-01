using System;
using System.IO;
using System.Linq;
using System.Reflection;

using Xunit;
using FluentAssertions;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Text;
using Transactions.Models;
using Transactions.Services;
using OfficeOpenXml.Table;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using ExcelClient.Tests.Builder;

namespace ExcelClient.Tests
{
    public class UnitTest1
    {
        [Fact]
        public void getSubCategoryModelFromExcel()
        {
            //var resourcePath = "ExcelClient.Tests.TestData.Transactions.xlxs";
            var resourcePath = "ExcelClient.Tests.TestData.Categories.xlsx";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);

            //var file = "Transactions.xlxs";
            var file = "Categories.xlsx";
            var savedFilePath = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file);
            var name = Path.GetFileNameWithoutExtension(savedFilePath);
            var fi = new FileInfo(savedFilePath);
            ExcelPackage ep = new ExcelPackage(new FileInfo(savedFilePath));

            ExcelPackage ep1 = new ExcelPackage(resourceAsStream);
            //ExcelPackage ep = new ExcelPackage(new FileInfo(resourcePath));
            ExcelWorksheet workSheet = ep.Workbook.Worksheets.FirstOrDefault();
            ExcelWorksheet workSheet1 = ep1.Workbook.Worksheets.FirstOrDefault();
            var json = ExcelServices.GetJsonFromTable(workSheet);
            var json1 = ExcelServices.GetJsonFromTable(workSheet1);
            var jarray = JArray.Parse(json1);
            List<SubCategory> subcategories = new List<SubCategory>();
            foreach (var subCategory in jarray)
            {
                subcategories.Add(JsonServices.GetSubCategory(subCategory));
            }

            var table = workSheet.Tables.FirstOrDefault();
            json.Should().NotBeNull();
        }
        [Fact]
        public void getSubMovementsModelFromExcel()
        {
            var resourceFileNema = "Transactions.xlxs";
            var resourcePath = $"ExcelClient.Tests.TestData.{resourceFileNema}";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);
            string filename;
            FileStream fs = resourceAsStream as FileStream;
            if (fs != null)
                filename = fs.Name;

            ExcelPackage ep = new ExcelPackage(resourceAsStream);

            ExcelWorksheet workSheet = ep.Workbook.Worksheets["Felles"];
            var json = ExcelServices.GetJsonFromTable(workSheet);
            var jarray = JArray.Parse(json);
            List<AccountMovement> acountsMovments = new List<AccountMovement>();
            foreach (var movment in jarray)
            {
                acountsMovments.Add(new ModelClassServices().JsonToAccountMovement(movment));
            }

            var table = workSheet.Tables.FirstOrDefault();
            json.Should().NotBeNull();
        }
        [Fact]
        public void ParseJsonToModelClassTest()
        {
            var resourceFileNema = "SubCategories.json";
            var resourcePath = $"ExcelClient.Tests.TestData.{resourceFileNema}";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);

            string json;
            JArray jArray;
            using (StreamReader r = new StreamReader(resourceAsStream))
            {
                json = r.ReadToEnd();
                jArray = JArray.Parse(json);
            }

            foreach (var item in jArray)
            {
                var noko = new ModelClassServices().JsonToSubCategory(item);
            }

            json.Should().NotBeNull();
        }
        [Fact]
        public void createMovementsExcel()
        {
            Stream SubCategoriesStream = GetAssemblyFile("Categories.xlsx");
            Stream AccountMovmentStream = GetAssemblyFile("Transactions.xlsx");

            ExcelWorksheet workSheet = ExcelServices.GetExcelWorksheet(AccountMovmentStream, "Felles");
            ExcelWorksheet workSheet2 = ExcelServices.GetExcelWorksheet(SubCategoriesStream);

            var subCategoriesjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet2));
            var accountMovmentjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet));
            List<SubCategory> subCategories = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);

            var modementsViewModels = ModelClassServices.CreateMovementsViewModels(accountMovements, subCategories, "Felles");

            var excelPkg = new ExcelPackage();
            try
            {

                ExcelWorksheet wsSheet = excelPkg.Workbook.Worksheets.Add("fellesResum");
                //Add Table Title
                ExcelServices.AddSheetHeading(wsSheet, "TableName");
                // Add "input" and "output" headet to Excel table
                //ExcelServices.AddTableHeadings(wsSheet, new[] { "col1", "col2", "col3" }, subCategoriesjArray.Count+ accountMovmentjArray.Count);
                //Add DMN Table to excel Sheet
                ExcelServices.CreateExcelTableFromMovementsViewModel(modementsViewModels, wsSheet, "TableName");

            }
            catch (Exception e)
            {
                var noko = e.Message;
            }
            var filename = "MovementsTests";
            var path = string.Concat(@"h:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));

            File.Exists(filePath).Should().BeTrue();
        }
        [Fact]
        public void createMonthYearSummaryExcel()
        {
            //Stream SubCategoriesStream = GetAssemblyFile("Categories.xlsx");
            //Stream AccountMovmentStream = GetAssemblyFile("Transactions.xlsx");
            ExcelWorksheet workSheet;
            ExcelWorksheet workSheet2;
            using (Stream AccountMovmentStream = GetAssemblyFile("Transactions.xlsx"))
            {
                workSheet = ExcelServices.GetExcelWorksheet(AccountMovmentStream, "Felles");
            }
            using (Stream SubCategoriesStream = GetAssemblyFile("Categories.xlsx"))
            {
                workSheet2 = ExcelServices.GetExcelWorksheet(SubCategoriesStream);
            }

            var subCategoriesjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet2));
            var accountMovmentjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet));
            List<SubCategory> categorisModel = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();
            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);

            var modementsViewModels = ModelClassServices.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            var excelPkg = new ExcelPackage();
            try
            {
                //Add excel sheet
                ExcelWorksheet wsSheet = excelPkg.Workbook.Worksheets.Add("MonthResume");

                //Add Table Title
                ExcelServices.AddSheetHeading(wsSheet, "TableName");

                // Add "input" and "output" headet to Excel table
                //ExcelServices.AddTableHeadings(wsSheet, new[] { "col1", "col2", "col3" }, subCategoriesjArray.Count+ accountMovmentjArray.Count);

                //Add transactions to excel Sheet
                ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, wsSheet, categoryList,0,null,true);
                //ExcelServices.CreateExcelTableFromMovementsViewModel(modementsViewModels, wsSheet, "TableName", categoryList);

            }
            catch (Exception e)
            {
                var noko = e.Message;
            }
            var filename = "MonthResumeTests";
            var path = string.Concat(@"h:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));

            File.Exists(filePath).Should().BeTrue();
        }

        [Fact]
        public void UpdateCashFlow()
        {
            JArray JsonmodementsViewModels;
            JArray JsonCategoryList;
            Encoding encoding = Encoding.GetEncoding(28591);

            using (StreamReader stream = new StreamReader(GetAssemblyFile("TransactionViewModelArray.json"), encoding, true))
            {
                JsonmodementsViewModels = JArray.Parse(stream.ReadToEnd());
            }
            using (StreamReader stream = new StreamReader(GetAssemblyFile("CategoriesArray.json"), encoding, true))
            {
                JsonCategoryList = JArray.Parse(stream.ReadToEnd());
            }
            var movementsViewModels = new List<MovementsViewModel>();
            foreach (var item in JsonmodementsViewModels)
            {
                movementsViewModels.Add(new ModelClassServices().JsonToMovementsViewModels(item));
            }
            List<string> categoryListTemp = new List<string>();
            foreach (var item in JsonCategoryList)
            {
                categoryListTemp.Add(item.ToString());
            }

            IEnumerable<string> categoryList = categoryListTemp;


            var excelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx"));
            try
            {
                var ExpensesWSheet = excelPkg.Workbook.Worksheets["Expenses details"];

                var year = 2018;

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
            var filename = "Budget Cashflow Temp";
            var path = string.Concat(@"h:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));

            File.Exists(filePath).Should().BeTrue();
        }

        [Fact]
        public void CreteCashFlowFromMovements()
        {
            var MovementsExcelPkg = new ExcelPackage(GetAssemblyFile("Movements.xlsx"));
            var WoorSheet = MovementsExcelPkg.Workbook.Worksheets.FirstOrDefault();
            var jsonFromTable = ExcelServices.GetJsonFromTable(WoorSheet);
            MovementsExcelPkg.Dispose();

            List<MovementsViewModel> movementsViewModels = JsonConvert.DeserializeObject<List<MovementsViewModel>>(jsonFromTable, JsonServices.GetJsonSerializerSettings());

            var categoryList = ModelClassServices.GetListOfCategories(movementsViewModels);

            //var movements = JsonConvert.DeserializeObject<List<MovementsViewModel>>(jsonFromTable, dateTimeConverter);
            var excelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow.xlsx"));
            try
            {
                var ExpensesWSheet = excelPkg.Workbook.Worksheets["Expenses details"];

                var year = 2018;

                //workSheet.Tables.Delete("YearExpenses");

                // add year expenses categoiers 
                ExcelServices.CreateYearExpensesTable(movementsViewModels, categoryList, year, ExpensesWSheet, "YearExpenses", "B38");

                // add year incoms categoiers 
                ExcelServices.CreateYearIncomsTable(movementsViewModels, categoryList, year, ExpensesWSheet, "YearIncoms", "B54");

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
            var filename = "Budget Cashflow Temp";
            var path = string.Concat(@"h:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));
            excelPkg.Dispose();

            File.Exists(filePath).Should().BeTrue();
        }

        [Fact]
        public void GetSubCategoryTest()
        {
            JArray JsonmodementsViewModels;
            JArray JsonCategoryList;
            Encoding encoding = Encoding.GetEncoding(28591);

            using (StreamReader stream = new StreamReader(GetAssemblyFile("TransactionViewModelArray.json"), encoding, true))
            {
                JsonmodementsViewModels = JArray.Parse(stream.ReadToEnd());
            }

            Stream SubCategoriesStream = GetAssemblyFile("Categories.xlsx");
            Stream AccountMovmentStream = GetAssemblyFile("Transactions.xlsx");

            ExcelWorksheet workSheet = ExcelServices.GetExcelWorksheet(AccountMovmentStream, "Felles");
            ExcelWorksheet workSheet2 = ExcelServices.GetExcelWorksheet(SubCategoriesStream);

            var subCategoriesjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet2));
            var accountMovmentjArray = JArray.Parse(ExcelServices.GetJsonFromTable(workSheet));
            List<SubCategory> subCategories = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);

            var Modelviews = new ModelViewBuilder().AddTextToMovemnt(new List<string> { "arg tur","arg tur", "argentina tur","tur argentina", "argentina tur" });

            var movementsViewModels = new List<MovementsViewModel>();
            foreach (var item in JsonmodementsViewModels)
            {
                movementsViewModels.Add(new ModelClassServices().JsonToMovementsViewModels(item));
            }

            List<MovementsViewModel> newList = new List<MovementsViewModel>();
            foreach (var movment in Modelviews)
            {
                newList.Add(ModelClassServices.UpdateMovementViewModelWithSubCategory(subCategories, movment));
            }

            newList.All(mv => !string.IsNullOrEmpty(mv.Category)).Should().BeTrue();
        }

        private static Stream GetAssemblyFile(string fileName)
        {
            var resourceFileNema = fileName;
            var resourcePath = $"ExcelClient.Tests.TestData.{resourceFileNema}";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);
            return resourceAsStream;
        }

        private static JArray GetJarrayfromJsonStream(Stream resourceAsStream)
        {
            string json;
            JArray jArray;
            List<SubCategory> subCategories = new List<SubCategory>();
            using (StreamReader r = new StreamReader(resourceAsStream, Encoding.Default, true))
            {
                json = r.ReadToEnd();
                jArray = JArray.Parse(json);
            }
            return jArray;
        }



    }
}
