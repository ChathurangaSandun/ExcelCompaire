using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ExcelCompaire.Enums;
using ExcelCompaire.Models;

namespace ExcelCompaire.ViewModels
{
    public class ExcelFileViewModels
    {
        public Dictionary<ExcelFileType, ExcelFile> ExcelFiles { get; set; }

    }
}