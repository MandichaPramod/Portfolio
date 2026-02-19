namespace PortfolioSheets.Models
{
    public class DividendRawRow
    {

        public string Symbol { get; set; }

        public DateTime Date { get; set; }

        public double Quantity { get; set; }

        public double DividendPerShare { get; set; }

        public double TotalDividend { get; set; }

        public double StockPrize { get; set; }
    }
}
