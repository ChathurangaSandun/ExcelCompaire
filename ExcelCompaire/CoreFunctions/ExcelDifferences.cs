using System;
using System.Collections.Generic;
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

        public ExcelDifferences(IXLWorksheet planWorksheet, IXLWorksheet productWorksheet)
        {
            _planWorksheet = planWorksheet;
            _productWorksheet = productWorksheet;

            List<SecondParamModel> planList = ManupilateSecondPart(_planWorksheet);
            List<SecondParamModel> productList = ManupilateSecondPart(_productWorksheet);

            IdentifyDeiff(planList, productList);
        }

        private void IdentifyDeiff(List<SecondParamModel> planList, List<SecondParamModel> productList)
        {
            var planDistinctItemCodes = planList.Select(o => o.ItemCode).Distinct().ToList();
            var productDistinctItemCodes = productList.Select(o => o.ItemCode).Distinct().ToList();

            foreach (string planDistinctItemCode in planDistinctItemCodes)
            {
                List<SecondParamModel> secondParamModels = planList.Where(i => i.ItemCode == planDistinctItemCode).ToList();

                





            }









        }


        public List<SecondParamModel> ManupilateSecondPart(IXLWorksheet worksheet)
        {
            System.Collections.Generic.List<SecondParamModel> secondParamModels = new List<SecondParamModel>();

            var lastRowNumber = worksheet.LastRowUsed().RowNumber();
            var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();
            var secontRangeAdditional = worksheet.Range(_startRow, 1, lastRowNumber, 1);

            var rowList = secontRangeAdditional.AsTable().DataRange.Rows().ToList();
            //.Select(o => o.FirstCell().Value).ToList()

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
                        paramModel.OrderQty = cellList[4].Value.ToString();
                        paramModel.PlanQty = cellList[5].Value.ToString();


                        paramModel.SMV = cellList[6].Value.ToString();
                        paramModel.SMO = cellList[7].Value.ToString();



                        List<string> dateList = new List<string>();
                        foreach (var value in cellList.Skip(8).Take(lastColumnNumber - 8).ToList())
                        {

                            if (value.Value.ToString() != "")
                            {
                                double d = Double.Parse(value.Value.ToString());
                                d = Math.Round(d, 0);

                                dateList.Add(d.ToString(CultureInfo.InvariantCulture).Split('.')[0]);

                            }
                            else
                            {
                                dateList.Add(value.Value.ToString());
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
}