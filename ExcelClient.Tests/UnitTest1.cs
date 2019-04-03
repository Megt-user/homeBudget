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
            var json = new ExcelServices().GetJson(workSheet);
            var json1 = new ExcelServices().GetJson(workSheet1);
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
            var json = new ExcelServices().GetJson(workSheet);
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
            JObject jObject;
            JArray jArray;
            using (StreamReader r = new StreamReader(resourceAsStream))
            {
                json = r.ReadToEnd();
                jArray = JArray.Parse(json);
                //jObject = JObject.Parse(json);
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

            var subCategoriesjArray = JArray.Parse(new ExcelServices().GetJson(workSheet2));
            var accountMovmentjArray = JArray.Parse(new ExcelServices().GetJson(workSheet));
            List<SubCategory> subCategories = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);

            var modementsViewModels = ModelClassServices.getListOfModementsViewModel(accountMovements, subCategories, "Felles");

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

            var subCategoriesjArray = JArray.Parse(new ExcelServices().GetJson(workSheet2));
            var accountMovmentjArray = JArray.Parse(new ExcelServices().GetJson(workSheet));
            List<SubCategory> categorisModel = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();
            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);

            var modementsViewModels = ModelClassServices.getListOfModementsViewModel(accountMovements, categorisModel, "Felles");

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
                ExcelServices.CreateExcelMonthSummaryTableFromMovementsViewModel(modementsViewModels, wsSheet, categoryList);
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

          
            var excelPkg = new ExcelPackage(GetAssemblyFile("Budget Cashflow - Copy.xlsx"));
            ExcelTable yearBudget;
            try
            {
             var  workSheet = excelPkg.Workbook.Worksheets["Expenses details"];
                yearBudget = workSheet.Tables.First(t => t.Name == "YearExpenses");
                //var tableStartAdress = yearBudget.Address.Start.Address;
                var tableStartAdress = "B54";
                var year = 2018;

                
                ExcelServices.AddYearExpensesTable(movementsViewModels, categoryList, workSheet, tableStartAdress, year);
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
            JObject jObject;
            string json;
            JArray jArray;
            List<SubCategory> subCategories = new List<SubCategory>();
            using (StreamReader r = new StreamReader(resourceAsStream, Encoding.Default, true))
            {
                json = r.ReadToEnd();
                jArray = JArray.Parse(json);
                //jObject = JObject.Parse(json);
            }
            return jArray;
        }

       
        
    }
}
