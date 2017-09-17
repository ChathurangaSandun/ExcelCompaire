using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using ExcelCompaire.Enums;

namespace ExcelCompaire.Models
{
    public class ExcelFile
    {
        public ExcelFileType ExcelFileType { get; set; }
       // public string FileName { get; set; }
        public string  FilePath { get; set; }
        public List<string> Worksheets { get; set; }
    }
}