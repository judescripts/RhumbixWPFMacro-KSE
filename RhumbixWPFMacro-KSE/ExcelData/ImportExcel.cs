using System;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class ImportExcel
    {
        public static Excel.Workbook OpenExcel()
        {

            var fileName = string.Empty;

            var openFileDialog = new OpenFileDialog
            {
                Filter = "csv files (*.csv) | *.csv |All files (*.*) |*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };
            if (openFileDialog.ShowDialog() == true)
            {
                fileName = openFileDialog.FileName;
            }
            try
            {
                var excelApp = new Excel.Application { Visible = true, DisplayAlerts = false };

                var workbook = excelApp.Workbooks.Open(fileName);

                return workbook;
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