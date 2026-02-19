namespace PortfolioSheets.Models
{
    public class Portfolio
    {
        public string Account { get; set; }
        public string Symbol { get; set; }
        public string Name { get; set; }
        public double Quantity { get; set; }
        public double BuyPrice { get; set; }
        public double CurrentPrice { get; set; }
        public double Dividend { get; set; }
        public double TotalCost { get; set; }
        public double CurrentValue { get; set; }
        public double PnL { get; set; }
        public double Change { get; set; }

        public List<CorpEvent> Events = new List<CorpEvent>();

    }
}
