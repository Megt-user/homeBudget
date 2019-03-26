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


namespace homeBudget.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    public class TransactionsController : Controller
    {

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Post(IFormFile transacation, IFormFile category)
        {

            long categoryFileSize = category.Length;

            var filePathTemp = Path.GetTempFileName();
            var filePath1 = Path.GetTempFileName();


            if (!IsFileValid(transacation) || !IsFileValid(category))
                return BadRequest();

            ExcelWorksheet transactionsWorkSheet;
            ExcelWorksheet categoriesWorkSheet;

            using (var stream = new FileStream(filePathTemp, FileMode.Create))
            {
                await transacation.CopyToAsync(stream);
                transactionsWorkSheet = ExcelServices.GetExcelWorksheet(stream, "Felles");
            }
            using (var stream = new FileStream(filePath1, FileMode.Create))
            {
                await category.CopyToAsync(stream);
                categoriesWorkSheet = ExcelServices.GetExcelWorksheet(stream);
            }

            var subCategoriesjArray = JArray.Parse(new ExcelServices().GetJson(categoriesWorkSheet));
            var accountMovmentjArray = JArray.Parse(new ExcelServices().GetJson(transactionsWorkSheet));

            List<AccountMovement> accountMovements = ModelClassServices.GetAccountMovmentsFromJarray(accountMovmentjArray);
            List<SubCategory> categorisModel = ModelClassServices.GetSubCategoriesFromJarray(subCategoriesjArray);
            IEnumerable<string> categoryList = categorisModel.Select(cat => cat.Category).Distinct();

            var modementsViewModels = ModelClassServices.getListOfModementsViewModel(accountMovements, categorisModel, "Felles");

            var excelPkg = new ExcelPackage();
            try
            {
                ExcelServices.CreateSheetWithTransactionMovments(modementsViewModels, excelPkg, "SheetName", "SheetHeading", "TableName");
            }
            catch (Exception e)
            {
                return BadRequest("Creating transaction sheet. Error message : " + e.Message);
            }
            try
            {
                ExcelServices.CreateSheetWithMonthSummary(modementsViewModels, excelPkg, "MonthSummaries", "This represent tables with Month summary by years", categoryList);
            }
            catch (Exception e)
            {
                return BadRequest("Creating MonthSummary Sheet. Error message : " + e.Message);
            }

            // Save Excel Package
            try
            {
                var filename = "MovementsTests";
                var path = Path.Combine(@"h:\", "Transactions");

                Directory.CreateDirectory(path);
                excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(filename, ".xlsx"))));

                return Ok(filename + "Created in:" + path);
            }
            catch
            {

                return BadRequest("Can't be saved");
            }
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