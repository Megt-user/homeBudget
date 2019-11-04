using System;
using System.Collections.Generic;
using System.Text;
using FluentAssertions;
using OfficeOpenXml;
using Xunit;

namespace homeBudget.Tests
{
    public class ExcelHelpersTests
    {
        [Fact]
        public void GetIndexFromColumnNameTest()
        {
            var streamFile = TestsHelper.GetAssemblyFile("Budget Cashflow.xlsx");

            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                var expensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];
                var table = expensesWSheet.Tables["Year_budget"];
                var noko = ExcelHelpers.GetAdressFromColumnName(table, "Bil");

            }
        }
        [Fact]
        public void GetNamesAdressTest()
        {
            var streamFile = TestsHelper.GetAssemblyFile("Budget Cashflow.xlsx");

            using (var cashflowExcelPkg = new ExcelPackage(streamFile))
            {
                var expensesWSheet = cashflowExcelPkg.Workbook.Worksheets["Expenses details"];
                var table = expensesWSheet.Tables["Year_budget"];
                var noko = ExcelHelpers.GetNamesAdress(TestsHelper.GetCategoryList(), table);
                noko["Familly"].Should().Be("I22");
            }
        }
    }
}
