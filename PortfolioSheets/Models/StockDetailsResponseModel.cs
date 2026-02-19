namespace PortfolioSheets.Models
{
    public class StockDetailsResponseModel
    {
        public List<BreakDownRow> stockTransactions;
        public List<RealisedPnLRow> stockTrades;
        public StockCMPSheet stockdetails;
        public List<DividendHistoryRow> dividendDetails;
        public List<DividendRawRow> dividendRawlist;
    }
}
