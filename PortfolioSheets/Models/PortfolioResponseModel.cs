namespace PortfolioSheets.Models
{
    public class PortfolioResponseModel
    {
        public Dictionary<string, List<BreakDownRow>> stockTransactions;
        public Dictionary<string, List<Portfolio>> AllLists;

        public List<MarketCapHistory> MarketCapHistory;

        public List<RealisedPnLRow> realisedPnLStocks;

        public List<BreakDownRow> tradeList;

    }

    public class ChartData
    {
        public string Label { get; set; } // Label for the chart (e.g., symbol or "Others")
        public double Value { get; set; }   // Corresponding value (PnL)
        public string Percent { get; set; }   // Corresponding value (PnL)
    }
}
