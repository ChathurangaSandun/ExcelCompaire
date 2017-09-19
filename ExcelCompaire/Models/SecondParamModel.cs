using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelCompaire.Models
{
    public class SecondParamModel
    {
        public string ItemCode { get; set; }
        public string Category { get; set; }
        public string  Style { get; set; }
        public string Item { get; set; }
        public string OrderQty { get; set; }
        public string PlanQty { get; set; }
        public string SMV{ get; set; }
        public string SMO { get; set; }

        public Dictionary<int,string> DateList { get; set; }




    }
}