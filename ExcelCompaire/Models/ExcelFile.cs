using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ClosedXML.Excel;

namespace ExcelCompaire.Models
{
    public class ExcelFile
    {

       // public string FileName { get; set; }
        public string  FilePath { get; set; }
        public XLWorkbook Workbook { get; set; }
    }
}