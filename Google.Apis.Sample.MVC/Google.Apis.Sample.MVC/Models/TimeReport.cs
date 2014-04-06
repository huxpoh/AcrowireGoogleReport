namespace Google.Apis.Sample.MVC.Models
{
    public class TimeReport
    {
        public string Date { set; get; }
        public string Project { set; get; }
        public string Task { set; get; }
        public string Type { set; get; }
        public string Description { set; get; }
        public decimal Duration { set; get; }
        public decimal Overtime { set; get; }
    }
}