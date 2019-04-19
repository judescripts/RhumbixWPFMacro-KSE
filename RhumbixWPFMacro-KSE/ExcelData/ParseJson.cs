using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class ParseJson
    {
        /// <summary>
        /// Work with imported excel workbook, sort and parse json blobs into C# List 
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns>C# List of KseJson Type</returns>
        public List<KseJson> ParseJsonBlob(Workbook workbook)
        {
            var uniqueList = new List<KseJson>();
            var allList = new List<KseJson>();

            try
            {
                // Get imported workbook and worksheet and get all the ranges
                var xlSheet = (Worksheet)workbook.Sheets[1];
                var iLastRow = xlSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                var oJsonRange = xlSheet.Range["Z2", "Z" + iLastRow];

                // Sort for optimization
                xlSheet.UsedRange.Select();
                xlSheet.Sort.SortFields.Clear();
                xlSheet.Sort.SortFields.Add((Range)xlSheet.UsedRange.Columns["Z"], XlSortOn.xlSortOnValues,
                    XlSortOrder.xlAscending, Type.Missing, XlSortDataOption.xlSortNormal);
                var sort = xlSheet.Sort;
                sort.SetRange(xlSheet.UsedRange);
                sort.Header = XlYesNoGuess.xlYes;
                sort.MatchCase = false;
                sort.Orientation = XlSortOrientation.xlSortColumns;
                sort.Apply();

                // Go through the Json blobs int he range and add unto to C# List as unique values by their Ids.
                foreach (Range cell in oJsonRange)
                {
                    if ((string)cell.Value == "[]" && cell.Value == null) continue;
                    var jsonString = (string)cell.Value;
                    var kdeJson = KseJson.FromJson(jsonString);

                    var employeeId = xlSheet.Range["A" + cell.Row, "A" + cell.Row].Value2;
                    foreach (var json in kdeJson)
                    {
                        allList.Add(json);

                        var schema = json.Schema;
                        if (schema != "Trade Change") continue;

                        json.EmployeeId = employeeId.ToString();

                        var startRow = cell.Row;
                        if (uniqueList.Any(item => item.Id == json.Id)) continue;
                        json.StartingRow = startRow;
                        uniqueList.Add(json);
                    }

                }

                // Add effective number of line items per shift extra entries
                foreach (var json in allList)
                {
                    var count = allList.Count(x => x.Id == json.Id);

                    foreach (var item in uniqueList)
                    {
                        if (item.Count < 1 && item.Id == json.Id)
                        {
                            item.Count = count;
                        }
                    }

                }

                return uniqueList;

            }
            catch (Exception ex)
            {
                using (var file = new System.IO.StreamWriter(@".\exceptionlog.txt"))
                {
                    file.WriteLine(ex.Message);
                }
            }

            return null;
        }
    }
}
