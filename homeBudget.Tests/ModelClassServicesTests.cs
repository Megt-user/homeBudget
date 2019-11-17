
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;
using FluentAssertions;
using System.Text;
using homeBudget.Models;
using homeBudget.Services;

namespace homeBudget.Tests
{
    public class ModelClassServicesTests
    {
        [Fact]
        public void Test1()
        {
            ExcelWorksheet workSheet;
            ExcelWorksheet workSheet2;
            using (Stream AccountMovmentStream = TestsHelper.GetAssemblyFile("Transactions.xlsx"))
            {
                workSheet = ExcelHelpers.GetExcelWorksheet(AccountMovmentStream, "Felles");
            }
            using (Stream SubCategoriesStream = TestsHelper.GetAssemblyFile("Categories.xlsx"))
            {
                workSheet2 = ExcelHelpers.GetExcelWorksheet(SubCategoriesStream);
            }

            var workSheet2Table = workSheet2.Tables.FirstOrDefault();
            var workSheetTable = workSheet.Tables.FirstOrDefault();
            var subCategoriesjArray = ExcelConverter.GetJsonFromTable(workSheet2Table);
            var accountMovmentjArray = ExcelConverter.GetJsonFromTable(workSheetTable);
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(accountMovmentjArray);

            var modementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            var jsonModementsViewModels = JArray.Parse(JsonConvert.SerializeObject(modementsViewModels));


            jsonModementsViewModels.Should().NotBeNullOrEmpty();
            //     var filename = "Budget Cashflow Temp";
            //var path = string.Concat(@"h:\temp\");
            //Directory.CreateDirectory(path);
            //var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            ////excelPkg?.SaveAs(new FileInfo(filePath));

            //File.Exists(filePath).Should().BeTrue();

        }

        [Fact]
        public void Test2()
        {
            //Get MovementsModel
            JArray JsonmodementsViewModels;
            Encoding encoding = Encoding.GetEncoding(28591);

            using (StreamReader stream = new StreamReader(TestsHelper.GetAssemblyFile("TransactionViewModelArray.json"), encoding, true))
            {
                JsonmodementsViewModels = JArray.Parse(stream.ReadToEnd());
            }

            var movementsViewModels = new List<TransactionViewModel>();
            foreach (var item in JsonmodementsViewModels)
            {
                movementsViewModels.Add(new ModelConverter().JsonToMovementsViewModels(item));
            }

            // Get Categories
            JArray JsonCategoryList;
            using (StreamReader stream = new StreamReader(TestsHelper.GetAssemblyFile("CategoriesArray.json"), encoding, true))
            {
                JsonCategoryList = JArray.Parse(stream.ReadToEnd());
            }
            List<string> categoryListTemp = new List<string>();
            foreach (var item in JsonCategoryList)
            {
                categoryListTemp.Add(item.ToString());
            }

            IEnumerable<string> categoryList = categoryListTemp;


            var noko = ModelOperation.AverageforCategory(movementsViewModels, "Mat",null, 6, true);
            var noko1 = ModelOperation.AverageforCategory(movementsViewModels, "Mat",2018,null, true);
            var noko2 = ModelOperation.AverageforCategory(movementsViewModels, "Mat",2018, 6, true);

            noko.Should().BeGreaterThan(0);

        }



       
    }
}
