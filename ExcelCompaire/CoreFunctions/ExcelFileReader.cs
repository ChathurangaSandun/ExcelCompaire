using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using ExcelCompaire.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelCompaire.CoreFunctions
{
    public class ExcelFileReader
    {
        private string _filePath;
        public ExcelFileReader(string filePath)
        {
            _filePath = filePath;
        }


        //public ExcelFile GetFileWorkSheets()
        //{
        //    string fileName = _filePath;
        //    var workbook = new XLWorkbook(fileName);

        //    List<string> xlWorksheets = workbook.Worksheets.Select(o => o.Name).ToList();

        //    return new ExcelFile()
        //    {
        //        WorkSheetList = xlWorksheets,
        //        FilePath = _filePath
        //    }; 
            

        //}

     




        public void ReadExcel()
        {
            try
            {
                string fileName = @"C:\Users\Chathuranga.Sandun\Desktop\chathu\1.xlsx";
                var workbook = new XLWorkbook(fileName);
                var ws1 = workbook.Worksheet(3);

                var lstNumber =  ws1.LastRowUsed().RowNumber();

                var range = ws1.Range(15, 1, lstNumber, 1);

                var list = range.AsTable().DataRange.Rows().ToList();
                //.Select(o => o.FirstCell().Value).ToList()

                var itemLastRow = list.Find(o =>
                {
                    var value = o.FirstCell().Value;
                    return value != null && value == "Clock hours";
                });

                int lastItemRowNumber = itemLastRow.RowNumber();


                var threePart = ws1.Range(15, 1, lastItemRowNumber, 8);



                








            }
            catch (Exception e)
            {
                throw;
            }
        }
    }
}