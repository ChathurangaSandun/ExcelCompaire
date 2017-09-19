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
        public List<List<CoupleOfSheet>> ExcelFileDataList { get; set; }
        public List<int> UnionDateList { get; set; }
        

        public ExcelDifferences(IXLWorksheet planWorksheet, IXLWorksheet productWorksheet)
        {
            _planWorksheet = planWorksheet;
            _productWorksheet = productWorksheet;

            List<SecondParamModel> planList = ManupilateSecondPart(_planWorksheet);
            List<SecondParamModel> productList = ManupilateSecondPart(_productWorksheet);

            UnionDateList = CreateDateList(_planWorksheet,_productWorksheet);

            List<List<CoupleOfSheet>> identifyDeiffList = IdentifyDeiff(planList, productList);

            ExcelFileDataList = identifyDeiffList;
        }

        public List<List<CoupleOfSheet>> GetAllExcelFileDataList()
        {
             return ExcelFileDataList;
        }

        public List<int> GetUnionDateList()
        {
            return UnionDateList;
        }

        private List<int> CreateDateList(IXLWorksheet planWorksheet, IXLWorksheet productWorksheet)
        {
            List<int> planDateList = DateList(_planWorksheet);
            List<int> productDateList = DateList(_productWorksheet);


            return planDateList.Union(productDateList).ToList().OrderBy(o => o).ToList();
        }

      
        private List<int> DateList(IXLWorksheet worksheet)
        {
             List<int> dateList = new List<int>();


            var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();

            //date list
            List<IXLRangeColumn> dateRangeList = worksheet.Range(13, 9, 13, lastColumnNumber).AsRange().Columns().ToList();


            foreach (var cell in dateRangeList)
            {
                var dateTime = cell.FirstCell().Value is DateTime ? (DateTime)cell.FirstCell().Value : new DateTime();
                var date = dateTime.Day;
                dateList.Add(date);
            }

            return dateList;
        }

        private List<List<CoupleOfSheet>> IdentifyDeiff(List<SecondParamModel> planList, List<SecondParamModel> productList)
        {
            var planDistinctItemCodes = planList.Select(o => o.ItemCode).Distinct().ToList();
            //var productDistinctItemCodes = productList.Select(o => o.ItemCode).Distinct().ToList();

           // int rowIndex = 1;


            List<List<CoupleOfSheet>> list = new List<List<CoupleOfSheet>>();



            foreach (var itemCode in planDistinctItemCodes)
            {


                var fullOuterJoin = GetValue(planList, productList, itemCode);

                list.Add(fullOuterJoin);
                //int rowIndexOut;


                //CreateOutPutExcel(fullOuterJoin, rowIndex, out rowIndexOut);

                //workSheet.Row(rowIndexOut-1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                //workSheet.Row(rowIndexOut-1).Style.Border.BottomBorderColor = XLColor.Black;

                //rowIndex = rowIndexOut + 1;
                
            }


            return list;
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


                List<int> worksheetDateList = DateList(worksheet);




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

                            paramModel.PlanQty = (d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                        }
                        else
                        {
                            paramModel.PlanQty = (pqty);
                        }


                        paramModel.SMV = cellList[6].Value.ToString();
                        paramModel.SMO = cellList[7].Value.ToString();


                        Dictionary<int , string> dateDictionary = new Dictionary<int, string>();


                        for (int j = 8; j < lastColumnNumber; j++)
                        {
                            IXLCell cell = cellList[j];

                            var dateValue = cell.Value.ToString();
                            if (dateValue!= "")
                            {

                                double d = Double.Parse(dateValue);
                                d = Math.Round(d, 0);

                                dateDictionary.Add(worksheetDateList[j-8],d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                            }
                            else
                            {
                                dateDictionary.Add(worksheetDateList[j - 8], dateValue);
                            }
                        }

                        paramModel.DateList = dateDictionary;

                        secondParamModels.Add(paramModel);
                    }
                    i++;
                }

            }
            return secondParamModels;

        }
    }

    public class CoupleOfSheet
    {
        public SecondParamModel PlanItem { get; set; }
        public SecondParamModel ProductItem { get; set; }

    }



}