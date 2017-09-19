using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using ExcelCompaire.Models;

namespace ExcelCompaire.CoreFunctions
{
    public class ExcelDifferences
    {
        private IXLWorksheet _planWorksheet;
        private IXLWorksheet _productWorksheet;
        private int _startRow = 15;
        private XLWorkbook workBook;
        private IXLWorksheet workSheet;

        public ExcelDifferences(IXLWorksheet planWorksheet, IXLWorksheet productWorksheet)
        {
            _planWorksheet = planWorksheet;
            _productWorksheet = productWorksheet;

            List<SecondParamModel> planList = ManupilateSecondPart(_planWorksheet);
            List<SecondParamModel> productList = ManupilateSecondPart(_productWorksheet);

            workBook = new XLWorkbook();
            workSheet = workBook.Worksheets.Add("Second Part");

            IdentifyDeiff(planList, productList);

        }

        private void IdentifyDeiff(List<SecondParamModel> planList, List<SecondParamModel> productList)
        {
            var planDistinctItemCodes = planList.Select(o => o.ItemCode).Distinct().ToList();
            //var productDistinctItemCodes = productList.Select(o => o.ItemCode).Distinct().ToList();

            int rowIndex = 1;




            foreach (var itemCode in planDistinctItemCodes)
            {


                var fullOuterJoin = GetValue(planList, productList, itemCode);


                int rowIndexOut;


                CreateOutPutExcel(fullOuterJoin, rowIndex, out rowIndexOut);

                workSheet.Row(rowIndexOut-1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                workSheet.Row(rowIndexOut-1).Style.Border.BottomBorderColor = XLColor.Black;

                rowIndex = rowIndexOut + 1;

            }
            
        }

        private void CreateOutPutExcel(List<CoupleOfSheet> itemCodeData, int rowIndex, out int rowIndexOut)
        {


            foreach (CoupleOfSheet data in itemCodeData)
            {
                int columnIndex = 0;
                workSheet.Cell(rowIndex, ++columnIndex).Value = "PLN";
                //plan
                SecondParamModel plan = data.PlanItem;

                if (plan != null)
                {
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.ItemCode;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.Category;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.Style;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.Item;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.OrderQty;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.PlanQty;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.SMO;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = plan.SMV;

                    for (int i = columnIndex; i < plan.DateList.Count; i++, columnIndex++)
                    {
                        workSheet.Cell(rowIndex, columnIndex).Value = plan.DateList[i];
                    }
                }
                else
                {
                   
                }

                rowIndex++;
                columnIndex = 0;
                workSheet.Cell(rowIndex, ++columnIndex).Value = "PRD";
                //plan
                SecondParamModel product = data.ProductItem;

                if (product != null)
                {
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.ItemCode;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.Category;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.Style;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.Item;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.OrderQty;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.PlanQty;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.SMO;
                    workSheet.Cell(rowIndex, ++columnIndex).Value = product.SMV;

                    for (int i = columnIndex; i < product.DateList.Count; i++, columnIndex++)
                    {
                        workSheet.Cell(rowIndex, columnIndex).Value = product.DateList[i];
                    }
                }
                else
                {
                   
                }


                rowIndex = rowIndex + 1;

            }

            workBook.SaveAs(@"C:\Users\Chathuranga.Sandun\Desktop\chathu\OutPutExcel.xlsx");
            rowIndexOut = rowIndex;



        }

        private static List<CoupleOfSheet> GetValue(List<SecondParamModel> planList, List<SecondParamModel> productList, string itemCode)
        {
            var leftJoin = from planItem in planList.Where(o => o.ItemCode == itemCode)
                           join productItem in productList.Where(o => o.ItemCode == itemCode)
                           on planItem.Style equals productItem.Style
                           into temp
                           from productItem in temp.DefaultIfEmpty()
                           select new
                           {
                               PlanItem = planItem,
                               ProductItem = productItem
                           };


            var rightJoin = from productItem in productList.Where(o => o.ItemCode == itemCode)
                            join planItem in planList.Where(o => o.ItemCode == itemCode)
                            on productItem.Style equals planItem.Style
                            into temp
                            from planItem in temp.DefaultIfEmpty()
                            select new
                            {
                                PlanItem = planItem,
                                ProductItem = productItem
                            };

            var fullOuterJoin = leftJoin.Union(rightJoin).ToList();

            List<CoupleOfSheet> coupleOfSheets = new List<CoupleOfSheet>();

            foreach (var item in fullOuterJoin)
            {
                coupleOfSheets.Add(new CoupleOfSheet()
                {
                    PlanItem = item.PlanItem,
                    ProductItem = item.ProductItem
                });
            }

            return coupleOfSheets;
        }


        public List<SecondParamModel> ManupilateSecondPart(IXLWorksheet worksheet)
        {
            
            System.Collections.Generic.List<SecondParamModel> secondParamModels = new List<SecondParamModel>();

            var lastRowNumber = worksheet.LastRowUsed().RowNumber();
            var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();
            var secontRangeAdditional = worksheet.Range(_startRow, 1, lastRowNumber, 1);


            var rowList = secontRangeAdditional.AsTable().DataRange.Rows().ToList();
            //.Select(o => o.FirstCell().Value).ToList()


            //date list
            List<IXLRangeColumn> dateRangeList = worksheet.Range(13, 9, 13, lastColumnNumber).AsRange().Columns().ToList();

            foreach (var cell in dateRangeList)
            {
                var dateTime = cell.FirstCell().Value is DateTime ? (DateTime) cell.FirstCell().Value : new DateTime();
                var date = dateTime.Day;
            }



            var itemLastRow = rowList.Find(o =>
            {
                var value = o.FirstCell().Value;
                return value != null && value is string && (string)value == "Clock hours";
            });

            if (itemLastRow != null)
            {
                int lastItemRowNumber = itemLastRow.RowNumber() - 3;

                var secondRange = worksheet.Range(_startRow, 1, lastItemRowNumber - 1, lastColumnNumber);

                var secondRowList = secondRange.AsRange().Rows().ToList();

                int i = 0;
                foreach (IXLRangeRow row in secondRowList)
                {
                    if (i % 2 == 0)
                    {

                        SecondParamModel paramModel = new SecondParamModel();
                        var cellList = row.Cells().ToList();
                        paramModel.ItemCode = cellList[0].Value.ToString();
                        paramModel.Category = cellList[1].Value.ToString();
                        paramModel.Style = cellList[2].Value.ToString();
                        paramModel.Item = cellList[3].Value.ToString();

                        //TODO : convert int string

                        //order
                        string oqty = cellList[4].Value.ToString();
                        if (oqty != "" && oqty != "TBA")
                        {

                            double d = Double.Parse(oqty);
                            d = Math.Round(d, 0);

                            paramModel.OrderQty = (d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                        }
                        else
                        {
                            paramModel.OrderQty = (oqty);
                        }


                        //plan
                        
                        string pqty = cellList[5].Value.ToString();
                        if (pqty != "" && pqty != "TBA")
                        {

                            double d = Double.Parse(pqty);
                            d = Math.Round(d, 0);

                            paramModel.OrderQty = (d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                        }
                        else
                        {
                            paramModel.OrderQty = (oqty);
                        }


                        paramModel.SMV = cellList[6].Value.ToString();
                        paramModel.SMO = cellList[7].Value.ToString();



                        List<string> dateList = new List<string>();
                        foreach (var value in cellList.Skip(8).Take(lastColumnNumber - 8).ToList())
                        {
                            var dateValue = value.Value.ToString();
                            if (value.Value.ToString() != "")
                            {

                                double d = Double.Parse(dateValue);
                                d = Math.Round(d, 0);

                                dateList.Add(d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                            }
                            else
                            {
                                dateList.Add(dateValue);
                            }
                        }

                        paramModel.DateList = dateList;

                        secondParamModels.Add(paramModel);
                    }
                    i++;
                }

            }
            return secondParamModels;

        }
    }

    class CoupleOfSheet
    {
        public SecondParamModel PlanItem { get; set; }
        public SecondParamModel ProductItem { get; set; }

    }



}