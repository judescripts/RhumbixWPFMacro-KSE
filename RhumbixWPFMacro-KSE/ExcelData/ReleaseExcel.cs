using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public class ReleaseExcel
    {
        public void ReleaseObject(Workbook workbook)
        {
            try
            {
                Marshal.ReleaseComObject(workbook);
                workbook = null;
            }
            catch (Exception ex)
            {
                workbook = null;
                using (var file = new System.IO.StreamWriter(@".\exceptionlog.txt"))
                {
                    file.WriteLine(ex.Message);
                }
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}