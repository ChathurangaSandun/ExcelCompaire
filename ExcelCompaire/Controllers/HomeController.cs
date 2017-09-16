using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelCompaire.CoreFunctions;
using ExcelCompaire.Enums;
using ExcelCompaire.Models;
using ExcelCompaire.ViewModels;
using Syncfusion.JavaScript;
using Syncfusion.JavaScript.Models;

namespace ExcelCompaire.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {


            return View();
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Import(ImportRequest importRequest)
        {
            return importRequest.SpreadsheetActions();
        }

        [HttpPost]
        public Dictionary<ExcelFileType, ExcelFile> UploadFiles(IEnumerable<HttpPostedFileBase> files)
        {
            Dictionary<ExcelFileType, ExcelFile> excelFiles = new Dictionary<ExcelFileType, ExcelFile>();



            if (ModelState.IsValid)
            {

                var i = 0;
                foreach (var file in files)
                {
                    if (file.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        if (fileName != null)
                        {
                            var path = Path.Combine(Server.MapPath("~/UploadFiles"), fileName);
                            file.SaveAs(path);


                            var excelFile = GetExcelWorkBook(path);

                            if (i == 0)
                            {
                                excelFiles[ExcelFileType.PlanExcel] = excelFile;
                            }
                            else
                            {
                                excelFiles[ExcelFileType.ProductExcel] = excelFile;
                            }
                            

                        }
                        i++;
                    }
                }
            }

            return excelFiles;
        }

        private ExcelFile GetExcelWorkBook(string path)
        {
            XLWorkbook workbook = new XLWorkbook(path);

            return new ExcelFile()
            {
                FilePath = path,
                Workbook = workbook
            };
        }
    }
}