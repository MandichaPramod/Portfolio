namespace PortfolioSheets.Models
{
    public class MarketCapHistory
    {

        public DateTime Month { get; set; }

        public MarketCapHistoryRow Invested { get; set; }

        public MarketCapHistoryRow Current { get; set; }

        public double StockMarketPercent { get; set; }

    }

    public class MarketCapHistoryRow
    {

        public double StockMarket { get; set; }

        public Dictionary<string,double> StockBrokerBalance { get; set; }

        public Dictionary<string, double> BankBalance { get; set; }

        public Dictionary<string, double> FD { get; set; }

        public Dictionary<string, double> Insurance { get; set; }
        public Dictionary<string, double> PF { get; set; }

        public Dictionary<string, double> Bond { get; set; }


        public double Total { get; set; }
    }
}
