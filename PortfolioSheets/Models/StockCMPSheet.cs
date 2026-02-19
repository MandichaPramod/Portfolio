namespace PortfolioSheets.Models
{
    public class StockCMPSheet
    {
        public string Symbol { get; set; }
        public string Name { get; set; }
        public double CMP { get; set; }
        public double OldCMP { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Volume { get; set; }
        public double MarketCap { get; set; }
        public double PriceToEarnings { get; set; }
        public double EarningsPerShare { get; set; }
        public double High52W { get; set; }
        public double Low52W { get; set; }
        public double Change { get; set; }
        public double ChangePercent { get; set; }
        public double CloseYest { get; set; }
        public double Shares { get; set; }
        public DateTime LastTradedTime { get; set; }
        public DateTime Lastupdate { get; set; }
    }
}
