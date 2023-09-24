using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkExcelAddins.Helpers
{
    public class MainHelper
    {
        public static void SaveAsPdf()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
           string sfileName_Document = workbook.Name;
            string sPath = workbook.Path;
            if (sPath == "" || sPath == null)
            {
                sPath = desktopFolder;
            }
            string sFullpath_pdf = sPath + "\\" + sfileName_Document + ".pdf";

            workbook.ExportAsFixedFormat(
                   Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                   sFullpath_pdf,
                   Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                   true,
                   true,
                   1,
                   10,
                   false);
        }
    }
}
