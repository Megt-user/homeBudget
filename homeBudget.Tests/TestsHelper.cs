using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using FluentAssertions;
using homeBudget.Models;
using homeBudget.Services;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace homeBudget.Tests
{
    public class TestsHelper
    {
        //Internal Methods
        private static string _resourcePath = "homeBudget.Tests.TestData.";
        private static string _mainresourcePath = "...homeBudget.Data.";
        public static Stream GetAssemblyFile(string fileName)
        {
            var resourceFileNema = fileName;
            var resourcePath = String.Concat(_resourcePath, fileName);
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);
            return resourceAsStream;
        } 
        public static Stream GetMainAssemblyFile(string fileName)
        {
            var resourceFileNema = fileName;
            var resourcePath = String.Concat(_mainresourcePath, fileName);
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);
            return resourceAsStream;
        }

        public static ExcelWorksheet GetExcelWorksheetFromFile(string fileName, string workSheetName = null)
        {
            ExcelWorksheet workSheet;
            using (Stream accountMovmentStream = GetAssemblyFile(fileName))
            {
                workSheet = ExcelHelpers.GetExcelWorksheet(accountMovmentStream, workSheetName);
            }
            return workSheet;
        }

        public static JArray GetJonsArrayFromFile(string fileName)
        {
            var encoding = Encoding.GetEncoding(28591);
            JArray jsonArray;
            using (StreamReader stream = new StreamReader(TestsHelper.GetAssemblyFile(fileName), encoding, true))
            {
                jsonArray = JArray.Parse(stream.ReadToEnd());
            }
            return jsonArray;
        }
        public static bool SaveExcrlPackage(ExcelPackage excelPackage, string fileName)
        {
            try
            {
                var path = string.Concat(@"C:\Transactions\");
                Directory.CreateDirectory(path);
                var filePath = Path.Combine(path, string.Concat(fileName, ".xlsx"));
                excelPackage?.SaveAs(new FileInfo(filePath));
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public static List<MovementsViewModel> GetMovementsViewModels()
        {
            List<Transaction> movementsModels = GetAccountMovements();
            movementsModels.Count.Should().Be(122);

            List<SubCategory> categorisModel = GetSubCategories();
            categorisModel.Count.Should().Be(105);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();


            var movementsViewModels = ModelConverter.CreateMovementsViewModels(movementsModels, categorisModel, "Felles");
            return movementsViewModels;
        }

        public static List<SubCategory> GetSubCategories()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            return ModelConverter.GetCategoriesFromJarray(jsonArray);
        }

        public static List<Transaction> GetAccountMovements()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            return ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
        }

        public static IEnumerable<string> GetCategoryList()
        {
            List<SubCategory> categorisModel = GetSubCategories();
            categorisModel.Count.Should().Be(105);
            return categorisModel.Select(cat => cat.Category).Distinct();
        }
    }
}
