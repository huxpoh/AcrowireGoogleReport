using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Google.Apis.Sample.MVC.Models
{
    public class TimeReportList
    {
        public TimeReportList()
        {
            GridList = new List<TimeReport>();
        }
        public List<TimeReport> GridList { set; get; }
    }
}