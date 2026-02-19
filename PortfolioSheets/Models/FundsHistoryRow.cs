namespace PortfolioSheets.Models
{
    public class FundsHistoryRow
    {

        public string Account { get; set; }

        public DateTime Date { get; set; }

        public double Debit { get; set; }

        public double Credit { get; set; }

        public string Remarks { get; set; }

    }
}
