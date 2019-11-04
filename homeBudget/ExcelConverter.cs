using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace homeBudget
{
    public class ExcelConverter
    {
        public static JArray GetJsonFromTable(ExcelTable table)
        {

            var jsonArray = new JArray();
            var dictionaryList = new List<Dictionary<string, string>>();
            var json = string.Empty;
            if (table != null)
            {
                var tableStartAdress = table.Address.Start.Address;
                var totalRows = table.Address.Rows;
                var totalColumns = table.Columns.Count;

                for (int row = 0 + 1; row < totalRows; row++)
                {
                    var valuesDictionary = new Dictionary<string, string>();

                    for (int column = 0; column < totalColumns; column++)
                    {
                        var objectName = table.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(tableStartAdress, 0, column)].Value;

                        var objectValue = table.WorkSheet.Cells[ExcelHelpers.AddRowAndColumnToCellAddress(tableStartAdress, row, column)].Value;

                        valuesDictionary.Add(objectName.ToString(), objectValue?.ToString());
                    }
                    dictionaryList.Add(valuesDictionary);
                }
                jsonArray = JArray.Parse(JsonConvert.SerializeObject(dictionaryList.ToArray()));
                json = Newtonsoft.Json.JsonConvert.SerializeObject(dictionaryList);
            }
            return jsonArray;
        }
      
    }
}
