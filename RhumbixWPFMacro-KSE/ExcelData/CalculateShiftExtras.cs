using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class CalculateShiftExtras
    {
        /// <summary>
        /// Validate Cost code and call append new line method
        /// </summary>
        /// <param name="uniqueList"></param>
        /// <param name="workbook"></param>
        public void ValidateCostCodes(List<KseJson> uniqueList, Workbook workbook)
        {
            var xlSheet = (Worksheet)workbook.Worksheets[1];
            var newEffectiveRow = 0;

            foreach (var item in uniqueList)
            {
                var effectiveRow = newEffectiveRow > 0 ? item.StartingRow + newEffectiveRow : item.StartingRow;
                var costCode = xlSheet.Range["I" + effectiveRow];
                var hours = xlSheet.Range["M" + effectiveRow].Value;

                var line = 1;
                while (line <= item.Count)
                {
                    // If Cost codes match
                    if (costCode.Value.ToString() == item.Store.CostCodeSelector.CodeCode)
                    {
                        // If pay type 1 hour is greater than trade hour
                        if (Convert.ToInt64(hours) >= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        {
                            xlSheet.Range["M" + effectiveRow].Value =
                                Convert.ToInt64(hours) - item.Store.HoursAsAboveTradeOnAboveCostCode;

                            var tradeHour = item.Store.HoursAsAboveTradeOnAboveCostCode;
                            newEffectiveRow = AppendNewLine(xlSheet, item, tradeHour, effectiveRow, null);
                            line = 4;
                        } // If pay type 1 hour is lesser than trade hour
                        else if (Convert.ToInt64(hours) <= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        { // If pay type 1 is empty
                            if (xlSheet.Range[$"K{effectiveRow}"].Value == null)
                            {
                                xlSheet.Range["K" + effectiveRow].Value = item.Store.TradeChange.Substring(0, 6);
                                var remainingHours = item.Store.HoursAsAboveTradeOnAboveCostCode - Convert.ToInt64(hours);
                                newEffectiveRow = AppendNewLine(xlSheet, item, null, effectiveRow, remainingHours);
                                line = 4;
                            } // If pay type 1 is used
                            else if (xlSheet.Range[$"K{effectiveRow + 1}"].Value == null)
                            {
                                xlSheet.Range[$"K{effectiveRow + 1}"].Value = item.Store.TradeChange.Substring(0, 6);
                                var remainingHours = item.Store.HoursAsAboveTradeOnAboveCostCode - Convert.ToInt64(hours);
                                newEffectiveRow = AppendNewLine(xlSheet, item, null, effectiveRow + 1, remainingHours);
                                line = 4;
                            } // If pay type 2 is used
                            else if (xlSheet.Range[$"K{effectiveRow + 2}"].Value == null)
                            {
                                xlSheet.Range[$"K{effectiveRow + 2}"].Value = item.Store.TradeChange.Substring(0, 6);
                                var remainingHours = item.Store.HoursAsAboveTradeOnAboveCostCode - Convert.ToInt64(hours);
                                newEffectiveRow = AppendNewLine(xlSheet, item, null, effectiveRow + 2, remainingHours);
                                line = 4;
                            }
                        }
                    } // If the first pay type doesn't match
                    else if (costCode.Offset[1, 0].Value.ToString() == item.Store.CostCodeSelector.CodeCode)
                    {
                        line += 1;
                        hours = xlSheet.Range["M" + effectiveRow + 1].Value;
                        // If pay type 1 hour is greater than trade hour
                        if (Convert.ToInt64(hours) >= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        {
                            xlSheet.Range["M" + effectiveRow + 1].Value =
                                Convert.ToInt64(hours) - item.Store.HoursAsAboveTradeOnAboveCostCode;

                            var tradeHour = item.Store.HoursAsAboveTradeOnAboveCostCode;
                            newEffectiveRow = AppendNewLine(xlSheet, item, tradeHour, effectiveRow + 1, null);
                            line = 4;
                        } // If pay type 1 hour is lesser than trade hour
                        else if (Convert.ToInt64(hours) <= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        { // If pay type 2 is empty
                            if (xlSheet.Range[$"K{effectiveRow + 1}"].Value == null)
                            {
                                xlSheet.Range["K" + effectiveRow + 2].Value = item.Store.TradeChange.Substring(0, 6);
                                var remainingHours = item.Store.HoursAsAboveTradeOnAboveCostCode - Convert.ToInt64(hours);
                                newEffectiveRow = AppendNewLine(xlSheet, item, null, effectiveRow + 2, remainingHours);
                                line = 4;
                            }
                            else
                            {
                                using (var file = new System.IO.StreamWriter(@".\exceptionlog.txt"))
                                {
                                    file.WriteLine($"Exceptions Error: Please double check Employee {item.EmployeeId}'s shift extra entry on row {item.StartingRow}");
                                }
                            }
                        }
                    } // If the second pay type doesn't match
                    else if (costCode.Offset[2, 0].Value.ToString() == item.Store.CostCodeSelector.CodeCode)
                    {
                        line += 2;
                        // If pay type 1 hour is greater than trade hour
                        if (Convert.ToInt64(hours) >= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        {
                            xlSheet.Range["M" + effectiveRow + 2].Value =
                                Convert.ToInt64(hours) - item.Store.HoursAsAboveTradeOnAboveCostCode;

                            var tradeHour = item.Store.HoursAsAboveTradeOnAboveCostCode;
                            newEffectiveRow = AppendNewLine(xlSheet, item, tradeHour, effectiveRow + 2, null);
                            line = 4;
                        } // If pay type 1 hour is lesser than trade hour
                        else if (Convert.ToInt64(hours) <= item.Store.HoursAsAboveTradeOnAboveCostCode)
                        {
                            // If pay type 2 is empty
                            if (xlSheet.Range[$"K{effectiveRow + 2}"].Value == null)
                            {
                                xlSheet.Range["K" + effectiveRow + 3].Value = item.Store.TradeChange.Substring(0, 6);
                                var remainingHours =
                                    item.Store.HoursAsAboveTradeOnAboveCostCode - Convert.ToInt64(hours);
                                newEffectiveRow = AppendNewLine(xlSheet, item, null, effectiveRow + 3, remainingHours);
                                line = 4;
                            }
                            else
                            {
                                using (var file = new System.IO.StreamWriter(@".\exceptionlog.txt"))
                                {
                                    file.WriteLine($"Exceptions Error: Please double check Employee {item.EmployeeId}'s shift extra entry on row {item.StartingRow}");
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Append new line and fill in the trade change hours
        /// </summary>
        /// <param name="xlSheet"></param>
        /// <param name="item"></param>
        /// <param name="tradeHours"></param>
        /// <param name="effectiveRow"></param>
        /// <param name="remainingHours"></param>
        public int AppendNewLine(Worksheet xlSheet, KseJson item, long? tradeHours, int effectiveRow, long? remainingHours = null)
        {
            var line = (Range)xlSheet.Rows[effectiveRow + 1];
            line.Insert();

            if (remainingHours == null)
            {
                // Employee Id
                xlSheet.Range["A" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["A" + effectiveRow].Value;
                // Job
                xlSheet.Range["D" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["D" + effectiveRow].Value;
                // Phase
                xlSheet.Range["H" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["H" + effectiveRow].Value;
                // Cost Code
                xlSheet.Range["I" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["I" + effectiveRow].Value;
                // Pay type
                xlSheet.Range["J" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["J" + effectiveRow].Value;
                // Pay group
                xlSheet.Range["K" + effectiveRow].Offset[1, 0].Value = item.Store.TradeChange.Substring(0, 6);
                // Hours
                xlSheet.Range["M" + effectiveRow].Offset[1, 0].Value = tradeHours;

                return 2;
            }
            else
            {
                // Employee Id
                xlSheet.Range["A" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["A" + effectiveRow].Value;
                // Job
                xlSheet.Range["D" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["D" + effectiveRow].Value;
                // Phase
                xlSheet.Range["H" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["H" + effectiveRow].Value;
                // Cost Code
                xlSheet.Range["I" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["I" + effectiveRow].Value;
                // Pay type
                xlSheet.Range["J" + effectiveRow].Offset[1, 0].Value = xlSheet.Range["J" + effectiveRow].Value;
                // Pay group
                xlSheet.Range["K" + effectiveRow].Offset[1, 0].Value = item.Store.TradeChange.Substring(0, 6);
                // Hours
                xlSheet.Range["M" + effectiveRow].Offset[1, 0].Value = remainingHours;

                return 2;
            }
        }
    }
}