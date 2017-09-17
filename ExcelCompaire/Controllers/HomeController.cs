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

        public static XLWorkbook PlanExcelFile;
        public static  XLWorkbook ProductExcelFile;



        public ActionResult Index()
        {

            string fileName = @"C:\Users\Chathuranga.Sandun\Desktop\chathu\1.xlsx";
            var workbook = new XLWorkbook(fileName);
            var plan = workbook.Worksheet("Discussion");


            string fileName2 = @"C:\Users\Chathuranga.Sandun\Desktop\chathu\2.xlsx";
            var workbook2 = new XLWorkbook(fileName2);
            var pro = workbook2.Worksheet("22-Aug");

            new ExcelDifferences(plan,pro);





            return View();
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Import(ImportRequest importRequest)
        {
            return importRequest.SpreadsheetActions();
        }

        [HttpPost]
        public ActionResult  UploadFiles(IEnumerable<HttpPostedFileBase> files)
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


                            var excelFile = GetExcelWorkBook(path,i);

                            if (i == 0)
                            {
                                excelFiles[ExcelFileType.PlanExcel] = excelFile;
                                //PlanExcelFile = excelFile;
                            }
                            else
                            {
                                excelFiles[ExcelFileType.ProductExcel] = excelFile;
                               // ProductExcelFile = excelFile;
                            }
                            

                        }
                        i++;
                    }
                }
            }


          
            return View( "Index",new ExcelFileViewModels()
            {
                ExcelFiles = excelFiles
            });
        }

        private ExcelFile GetExcelWorkBook(string path ,int i)
        {
            XLWorkbook workbook = new XLWorkbook(path);

            if (i == 0)
            {
                PlanExcelFile = workbook;
            }
            else if(i == 1)
            {
                ProductExcelFile = workbook;
            }

            return new ExcelFile()
            {
                FilePath = path,
                Worksheets = workbook.Worksheets.Select(o => o.Name).ToList()
            };
        }


        [HttpPost]
        public ActionResult SelectWorkSheet(string PlanExcel, string ProductExcel)
        {
            IXLWorksheet planWorksheet = PlanExcelFile.Worksheet(PlanExcel);
            IXLWorksheet productWorksheet = ProductExcelFile.Worksheet(ProductExcel);


            TempData["planWorksheet"] = planWorksheet;
            TempData["productWorksheet"] = productWorksheet;



            return RedirectToAction("Index", "DifferenceExcel");
        }

       
    }
}