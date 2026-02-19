using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class PortfolioController : Controller
    {
        [HttpGet("Portfolio/{portfolioName}")]
        public IActionResult Index([FromRoute] string portfolioName)
        {
            if(portfolioName == null)
                return RedirectToAction("Index","Home");

            //Get all Buy Transactions
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list = GoogleSheets.AllBreakdownSheetData("Breakdown");
            list.Sort((x, y) => x.Date.CompareTo(y.Date));
            var allStockTransactions = list.Where(x => x.Portfolio == portfolioName).ToList();

            var yearlyStocks = new List<BreakDownRow>();
            var monthlyStocks = new List<BreakDownRow>();

            foreach (var items in allStockTransactions)
            {

                if (yearlyStocks.Any(x => x.Date.Year == items.Date.Year))
                {
                    var pf = yearlyStocks.First(x => x.Date.Year == items.Date.Year);


                    pf.CurrentValue = pf.CurrentValue + items.CurrentValue;
                    pf.Dividend = pf.Dividend + items.Dividend;
                    pf.TotalCost = pf.TotalCost + items.TotalCost;
                    pf.DividendReinvestment = pf.DividendReinvestment + items.DividendReinvestment;
                }
                else
                {
                    yearlyStocks.Add(new BreakDownRow()
                    {
                        Date = items.Date,
                        CurrentValue = items.CurrentValue,
                        Dividend = items.Dividend,
                        TotalCost = items.TotalCost,
                        DividendReinvestment = items.DividendReinvestment,
                    });
                }
                if (monthlyStocks.Any(x => x.Date.Year == items.Date.Year && x.Date.Month == items.Date.Month))
                {
                    var pf = monthlyStocks.First(x => x.Date.Year == items.Date.Year && x.Date.Month == items.Date.Month);


                    pf.CurrentValue = pf.CurrentValue + items.CurrentValue;
                    pf.Dividend = pf.Dividend + items.Dividend;
                    pf.TotalCost = pf.TotalCost + items.TotalCost;
                    pf.DividendReinvestment = pf.DividendReinvestment + items.DividendReinvestment;
                }
                else
                {
                    monthlyStocks.Add(new BreakDownRow()
                    {
                        Date = items.Date,
                        CurrentValue = items.CurrentValue,
                        Dividend = items.Dividend,
                        TotalCost = items.TotalCost,
                        DividendReinvestment = items.DividendReinvestment,
                    });
                }
            }

            return View(new StockDetailsResponseModel()
            {
                stockTransactions = allStockTransactions
            });
        }
    }
}
