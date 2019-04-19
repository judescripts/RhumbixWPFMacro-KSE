using System;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class MainMacros
    {
        public void RunMacros()
        {
            var workbook = ImportExcel.OpenExcel();
            var json = new ParseJson();
            var uniqueList = json.ParseJsonBlob(workbook);
            var calculate = new CalculateShiftExtras();
            calculate.ValidateCostCodes(uniqueList, workbook);
            var cleanUp = new CleanUpFormat();
            cleanUp.RemoveAllJsonBlobs(workbook);
            cleanUp.SaveFileAs(workbook);
            var release = new ReleaseExcel();
            release.ReleaseObject(workbook);
        }
    }
}
