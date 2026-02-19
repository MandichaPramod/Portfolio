namespace PortfolioSheets.Models
{
    public class ExternalCorporateEvent
    {
        public string Symbol { get; set; }

        public string Name { get; set; }
        public string Segment { get; set; }
        public string Type { get; set; }
        public string Remarks { get; set; }
        public string Status { get; set; }
        public DateTime BeforeDate { get; set; }
        public DateTime ExDate { get; set; }
        public string Ratio { get; set; }
        public string Cost { get; set; }

    }

    public class CorpEvent
    {
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
