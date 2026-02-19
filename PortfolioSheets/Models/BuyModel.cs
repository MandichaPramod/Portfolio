namespace PortfolioSheets.Models
{
    public class BuyModel
    {
        public string Symbol { get; set; }
        public string Account { get; set; }
        public double Qty { get; set; }
        public double BuyPrice { get; set; }
        public DateTime BuyDate { get; set; }
        public string Portfolio { get; set; }
    }
}
