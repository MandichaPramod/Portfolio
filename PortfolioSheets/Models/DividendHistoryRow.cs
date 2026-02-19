namespace PortfolioSheets.Models
{
    public class DividendHistoryRow
    {

        public string Symbol { get; set; }

        public string Company { get; set; }

        public DateTime Date { get; set; }

        public double Dividend { get; set; }

        public string Remarks { get; set; }

    }
}
