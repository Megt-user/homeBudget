﻿using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using homeBudget.Models;
using OfficeOpenXml.Table;
using System.Drawing;
using System.Globalization;
using homeBudget.Services;

namespace ExcelClient
{
    public class ExcelServices
    {
        private static int _startRow = 3;
        private static int _startColumn = 1;
        private static string _heding = "B1";
        private static string _titlecell = "C1";

        public string GetJson(ExcelWorksheet ws)
        {
            var table = ws.Tables.FirstOrDefault();
            string json = string.Empty;
            if (table != null)
            {
                var TableStartRow = table.Address.Start.Row;
                var TableEndtRow = table.Address.End.Row;
                var dictionaryList = new List<Dictionary<string, string>>();

                for (int i = TableStartRow + 1; i <= TableEndtRow; i++)
                {
                    var valuesDictionary = new Dictionary<string, string>();
                    for (int j = table.Address.Start.Column; j <= table.Address.End.Column; j++)
                    {
                        //var headingPosition = string.Concat(GetColumnName(j), TableStartRow);
                        var cellTitle = ws.Cells[string.Concat(GetColumnName(j), TableStartRow)].Value;
                        //var cellName = string.Concat(GetColumnName(j), i);
                        var cellValue = ws.Cells[string.Concat(GetColumnName(j), i)].Value;
                        if (j == table.Address.Start.Column)
                            valuesDictionary.Add("Id", i.ToString());

                        if (valuesDictionary.ContainsKey(cellTitle.ToString()))
                        {
                            valuesDictionary.Add($"{cellTitle}_{string.Concat(GetColumnName(j), i)}", cellValue?.ToString());
                        }
                        else
                        {
                            valuesDictionary.Add(cellTitle.ToString(), cellValue?.ToString());
                        }
                    }
                    dictionaryList.Add(valuesDictionary);
                }
                json = Newtonsoft.Json.JsonConvert.SerializeObject(dictionaryList);
            }
            return json;
        }

        public static void AddSheetHeading(ExcelWorksheet wsSheet, string tableTitle)
        {
            wsSheet.Cells[_heding].Value = "Table Name";
            wsSheet.Cells[_titlecell].Value = tableTitle;
            wsSheet.Cells[_heding].Style.Font.Size = 12;
            wsSheet.Cells[_heding].Style.Font.Bold = true;
            wsSheet.Cells[_heding].Style.Font.Italic = true;
        }
        public static void AddTableHeadings(ExcelWorksheet wsSheet, string[] columnsNames, int rows)
        {
            using (ExcelRange rng = wsSheet.Cells[_startRow, _startColumn, _startRow, columnsNames.Length - 1])
            {
                tableHeadingFormat(rng, "Input");
            }
        }

        public static void CreateExcelTableFromMovementsViewModel(List<MovementsViewModel> movementsModel, ExcelWorksheet wsSheet, string tableName)
        {
            var excelColumns = MovementsViewModel.excelColumns;

            // Calculate size of the table
            var endRow = _startRow + movementsModel.Count + 1;
            var endColum = _startColumn + excelColumns.Count;

            // Create Excel table Header
            using (ExcelRange rng = wsSheet.Cells[_startRow, _startColumn, endRow, endColum])
            {
                //Indirectly access ExcelTableCollection class
                ExcelTable table = wsSheet.Tables.Add(rng, tableName);
                var color = Color.FromArgb(250, 199, 111);
                //Set Columns position & name
                var i = 0;
               foreach (var property in excelColumns)
                {
                    table.Columns[i].Name = property;
                    i++;
                }

                // Add empty cell for annotation
                //table.Columns[i].Name = "Annotation";
                //AddExcelCellByRowAndColumn(stratColumn + i, stratRow + 1, " ", wsSheet, color);

                table.ShowHeader = true;
                table.ShowFilter = true;
                //table.ShowTotal = true;
            }

            // Set Excel table content
            var tableStartColumn = _startColumn;
            var row = _startRow + 1;
            foreach (var movement in movementsModel)
            {
                // account Movements
                foreach (var propertyName in excelColumns)
                {
                    //Get Property name value
                    var propertyValue =ModelClassServices.GetPropertyValue(movement,propertyName);
                    //add value tu excel cell
                    AddExcelCellByRowAndColumn(tableStartColumn, row, propertyValue?.ToString(), wsSheet);
                    tableStartColumn++;
                }
                tableStartColumn = _startColumn;
                row++;
            }

            //var noko = dict2.Keys.Except(dict.Keys);
            //var noko2 = dict.Keys.Except(dict2.Keys);
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        private static SubCategory ChoseSubCategory(IEnumerable<SubCategory> subcategoriesMatch)
        {
            var subcategory = new SubCategory();
            var moreThanOneCategory = subcategoriesMatch.Select(sub => sub.Category).Distinct().Count() > 1;
            if (moreThanOneCategory)
            {
                var subcategoryNames = subcategoriesMatch.Select(sub => sub.KeyWord).ToArray();
                var subcategoryCategories = subcategoriesMatch.Select(sub => sub.Category).ToArray();
                subcategory.KeyWord = string.Join(",", subcategoryNames);
                subcategory.Category = string.Join(",", subcategoryCategories);
                subcategory.SupPorject = "Mismatch";
            }
            else
            {
                subcategory = subcategoriesMatch.FirstOrDefault();
            }
            return subcategory;
        }

        private static void AddSubcategoryToExcel(ExcelWorksheet wsSheet, List<string> subCategoryProperties, int subCategoryColumn, int row, SubCategory subcategory)
        {
            foreach (var property in subCategoryProperties)
            {
                string valueString = valueString = subcategory.GetType().GetProperty(property).GetValue(subcategory, null)?.ToString();
                AddExcelCellByRowAndColumn(subCategoryColumn, row, valueString, wsSheet);
                subCategoryColumn++;
            }
        }

        // internal method
        private static void tableHeadingFormat(ExcelRange excelRange, string text)
        {
            excelRange.Value = text;
            excelRange.Style.Font.Size = 12;
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Italic = true;
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelRange.Merge = true;
            excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }
        private static void AddExcelCellByRowAndColumn(int column, int row, string value, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellName = string.Concat(GetColumnName(column), row);
            using (ExcelRange rng1 = wsSheet.Cells[cellName])
            {
                if (color.HasValue)
                {
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(color.Value);
                }
                rng1.Value = value;
            }
        }

        private int GetColumnIndex(string columnName)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var index = letters.ToLower().IndexOf(columnName.ToLower());

            return index + 1;
        }

        private static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];
            else
                value += letters[index % letters.Length - 1];

            return value;
        }


    }



}

