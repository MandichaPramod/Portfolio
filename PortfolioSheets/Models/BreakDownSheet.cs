namespace PortfolioSheets.Models
{
    public class BreakDownRow
    {
        public string Symbol { get; set; }

        public string Name { get; set; }
        public string Account { get; set; }
        public DateTime Date { get; set; }
        public double Quantity { get; set; }
        public double BuyPrice { get; set; }
        public double TotalCost { get; set; }
        public double CurrentValue { get; set; }
        public double Dividend { get; set; }
        public double DividendReinvestment { get; set; }


        public string Portfolio { get; set; }
        public string Tag { get; set; }
        public int Ageing { get; set; }
        public double Change { get; set; }

    }
}
