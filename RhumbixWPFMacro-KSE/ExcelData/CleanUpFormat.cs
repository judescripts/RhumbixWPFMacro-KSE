using Microsoft.Office.Interop.Excel;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class CleanUpFormat
    {
        public void RemoveAllJsonBlobs(Workbook workbook)
        {
            var xlSheet = (Worksheet)workbook.Worksheets[1];
            var iLastRow = xlSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            var oJsonRange = xlSheet.Range["Z1", $"Z{iLastRow}"];
            oJsonRange.ClearContents();
        }

        public void SaveFileAs(Workbook workbook)
        {
            var save = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV|*.csv",
                Title = "KSE Export v2",
                CheckPathExists = true,
                FileName = "KSE Export v2"
            };
            if (save.ShowDialog() == true)
            {
                workbook.SaveAs(save.FileName);
            }
        }
    }
}
