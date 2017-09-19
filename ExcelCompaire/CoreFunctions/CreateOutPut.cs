using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelCompaire.Models;

namespace ExcelCompaire.CoreFunctions
{
    public class CreateOutPut
    {
        private XLWorkbook workbook;
        private IXLWorksheet worksheet;
        public CreateOutPut()
        {
            
            workbook = new XLWorkbook();
            worksheet = workbook.Worksheets.Add("Second Part");





        }

        public void CreateOutPutExcel(List<List<CoupleOfSheet>> allList, List<int> unionDateList)
        {

            DrawHeaders(worksheet,unionDateList);



            //int columnIndex;
            int rowIndex = 2;
            foreach (List<CoupleOfSheet> coupleOfSheets in allList)
            {
                foreach (CoupleOfSheet coupleOfSheet in coupleOfSheets)
                {
                    worksheet.Cell(rowIndex, 1).Value = "PLN";
                    worksheet.Cell(rowIndex+1, 1).Value = "PRD";


                    int column = 2;
                    var plan = coupleOfSheet.PlanItem;
                    var product = coupleOfSheet.ProductItem;
                    if (coupleOfSheet.PlanItem != null && coupleOfSheet.ProductItem != null)
                    {
                        worksheet.Cell(rowIndex, column).Value = plan.ItemCode;
                        worksheet.Cell(rowIndex + 1, column).Value = product.ItemCode;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Category;
                        worksheet.Cell(rowIndex+1, column).Value = product.Category;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Style;
                        worksheet.Cell(rowIndex + 1, column).Value = product.Style;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Item;
                        worksheet.Cell(rowIndex + 1, column).Value = product.Item;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.OrderQty;
                        worksheet.Cell(rowIndex + 1, column).Value = product.OrderQty;
                        if (plan.OrderQty != product.OrderQty)
                        {
                            worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                            worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        }
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.PlanQty;
                        worksheet.Cell(rowIndex + 1, column).Value = product.PlanQty;
                        if (plan.PlanQty != product.PlanQty)
                        {
                            worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                            worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        }
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.SMV;
                        worksheet.Cell(rowIndex + 1, column).Value = product.SMV;
                        if (plan.SMV != product.SMV)
                        {
                            worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                            worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        }
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.SMO;
                        worksheet.Cell(rowIndex + 1, column).Value = product.SMO;
                        if (plan.SMO != product.SMO)
                        {
                            worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                            worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        }

                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;


                        foreach (var date in unionDateList)
                        {
                            bool hasPlanDate = plan.DateList.ContainsKey(date);
                            bool  hasProductDate = product.DateList.ContainsKey(date);
                            column++;
                            if (hasPlanDate)
                            {
                                worksheet.Cell(rowIndex, column).Value = plan.DateList[date];
                            }
                            if (hasProductDate)
                            {
                                worksheet.Cell(rowIndex+1, column).Value = product.DateList[date];
                            }

                            if (hasProductDate && hasPlanDate)
                            {
                                if (plan.DateList[date] != product.DateList[date])
                                {
                                    worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                    worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                }
                            }
                            else if (hasProductDate || hasPlanDate)
                            {
                                if (hasProductDate)
                                {
                                    if (product.DateList[date] != "")
                                    {
                                        worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                        worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                    } 
                                }else
                                {
                                    if (plan.DateList[date] != "")
                                    {
                                        worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                        worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                    }
                                }



                            }
                           
                        }
                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                        
                    }
                    else if (coupleOfSheet.PlanItem != null && coupleOfSheet.ProductItem == null)
                    {
                        worksheet.Cell(rowIndex, column).Value = plan.ItemCode;
                        worksheet.Cell(rowIndex + 1, column).Value = plan.ItemCode;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Category;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Style;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.Item;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.OrderQty;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.PlanQty;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.SMV;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";
                        column++;
                        worksheet.Cell(rowIndex, column).Value = plan.SMO;
                        worksheet.Cell(rowIndex + 1, column).Value = "-";

                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;


                        foreach (var date in unionDateList)
                        {
                            bool hasPlanDate = plan.DateList.ContainsKey(date);
                           
                            column++;
                            if (hasPlanDate)
                            {
                                worksheet.Cell(rowIndex, column).Value = plan.DateList[date];
                                if (plan.DateList[date] != "")
                                {
                                    worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                    worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                }
                            }
                          
                        }
                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;

                        worksheet.Cell(rowIndex, 2).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        worksheet.Cell(rowIndex + 1, 2).Style.Fill.BackgroundColor = XLColor.RedMunsell;

                    }
                    else
                    {
                        worksheet.Cell(rowIndex, column).Value = product.ItemCode.ToString();
                        worksheet.Cell(rowIndex + 1, column).Value = product.ItemCode.ToString();
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex+1, column).Value = product.Category;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.Style;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.Item;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.OrderQty;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.PlanQty;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.SMV;
                        column++;
                        worksheet.Cell(rowIndex, column).Value = "-";
                        worksheet.Cell(rowIndex + 1, column).Value = product.SMO;


                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;


                        foreach (var date in unionDateList)
                        {
                            bool hasProductDate = product.DateList.ContainsKey(date);
                            column++;
                           
                            if (hasProductDate)
                            {
                                worksheet.Cell(rowIndex + 1, column).Value = product.DateList[date];

                                if (product.DateList[date] != "")
                                {
                                    worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                    worksheet.Cell(rowIndex + 1, column).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                                }
                               
                            }

                            
                        }
                        worksheet.Column(column).Style.Border.RightBorder = XLBorderStyleValues.Thick;

                        worksheet.Cell(rowIndex, 2).Style.Fill.BackgroundColor = XLColor.RedMunsell;
                        worksheet.Cell(rowIndex + 1, 2).Style.Fill.BackgroundColor = XLColor.RedMunsell;

                    }

                    worksheet.Row(rowIndex + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    rowIndex += 2;

                }
                worksheet.Row(rowIndex-1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                rowIndex += 1;
            }


            workbook.SaveAs(@"C:\Users\Chathuranga.Sandun\Desktop\chathu\OutPutExcel.xlsx");
        }

        private void DrawHeaders(IXLWorksheet xlWorksheet, List<int> unionDateList)
        {
            int column = 0;

            xlWorksheet.Cell(1, ++column).Value = "";
            xlWorksheet.Cell(1, ++column).Value = "ItemCode";
            xlWorksheet.Cell(1, ++column).Value = "Category";
            xlWorksheet.Cell(1, ++column).Value = "Style";
            xlWorksheet.Cell(1, ++column).Value = "Item";
            xlWorksheet.Cell(1, ++column).Value = "Oder Qty";
            xlWorksheet.Cell(1, ++column).Value = "Plan Qty";
            xlWorksheet.Cell(1, ++column).Value = "SMV";
            xlWorksheet.Cell(1, ++column).Value = "SMO";

            foreach (var date in unionDateList)
            {

                xlWorksheet.Cell(1, ++column).Value = date.ToString();
            }

            worksheet.Row(1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;


        }
    }
}