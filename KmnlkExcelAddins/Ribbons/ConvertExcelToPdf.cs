using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
//----< Excel Addin >----
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using KmnlkExcelAddins.Helpers;
//----</ Excel Addin >----
namespace KmnlkExcelAddins.Ribbons
{
    public partial class ConvertExcelToPdf
    {
        private void ConvertExcelToPdf_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_SaveAsPdf_Click(object sender, RibbonControlEventArgs e)
        {
            MainHelper.SaveAsPdf();
        }
    }
}
