namespace PortfolioSheets.Models
{
    public class RealisedPnLRow
    {
        public string Account { get; set; }

        public string Symbol { get; set; }

        public string Name { get; set; }

        public DateTime SellDate { get; set; }

        public double Quantity { get; set; }

        public double BuyPrice { get; set; }

        public double SellPrice { get; set; }

        public double Dividend { get; set; }

        public double TotalCost { get; set; }

        public double SellValue { get; set; }

        public double PnL { get; set; }

        public double PnLAfterCharges { get; set; }

        public double TotalCharges { get; set; }

        public double Brokerage { get; set; }

        public double STT { get; set; }

        public double TranCharges { get; set; }

        public double GST { get; set; }

        public double SEBICharges { get; set; }

        public double StampDuty { get; set; }

        public double OtherCharges { get; set; }

        public string Type { get; set; }

    }
}
