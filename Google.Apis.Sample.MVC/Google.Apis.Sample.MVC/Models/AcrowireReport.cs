using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Google.Apis.Sample.MVC.Models
{
    public class AcrowireReport
    {
        public string Date { set; get; }
        public string Person { set; get; }
        public string Project { set; get; }
        public string Summary { set; get; }
        public string BillingStatus { set; get; }
        public string Hours { set; get; }
    }

    public class AcrowireReportList
    {
        public AcrowireReportList()
        {
            GridList = new List<AcrowireReport>();
        }
        public List<AcrowireReport> GridList { set; get; }
    }
}