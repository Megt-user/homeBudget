using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentAssertions;
using homeBudget.Models;
using homeBudget.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Xunit;

namespace homeBudget.Tests
{
    public class ExcelConverterTests
    {

        [Fact]
        public void UpdateClassesTableValues()
        {
            var streamFile = TestsHelper.GetAssemblyFile("Transactions Update With Categories.xlsx");
            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                IEnumerable<string> categoryList = TestsHelper.GetCategoryList();
                var expensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Transactions"];

                var transactions = expensesWSheet.Tables.FirstOrDefault();
                var jsonArray = ExcelConverter.GetJsonFromTable(transactions);
                var categoriesAverageWorkSheet = cashflowExcelPkg.Workbook.Worksheets["Categories Average"];
                var categoriesAverageWorkSheet1 = cashflowExcelPkg.Workbook.Worksheets["Categories Average1"];
                if (categoriesAverageWorkSheet != null)
                {

                }
                jsonArray.Count.Should().Be(193);

                var noko = jsonArray.ToObject<List<TransactionViewModel>>();
                List<TransactionViewModel> movementsViewModels = JsonConvert.DeserializeObject<List<TransactionViewModel>>(jsonArray.ToString(), JsonServices.GetJsonSerializerSettings());
                movementsViewModels.Count.Should().Be(193);
            }
        }

    }
}
