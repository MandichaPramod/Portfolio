using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class StockController : Controller
    {
        [HttpGet("Stock/{symbol}")]
        public IActionResult Index([FromRoute] string symbol)
        {
            if(symbol == null)
                return RedirectToAction("Index","Home");

            //Get all Buy Transactions
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list = GoogleSheets.AllBreakdownSheetData("Breakdown");
            list.Sort((x, y) => x.Date.CompareTo(y.Date));
            var allStockTransactions = list.Where(x => x.Symbol == symbol).ToList();

            //Get All Positional Trades
            var pnllist = GoogleSheets.GetRealisedPnL("InActive");
            var stockslist = GoogleSheets.GetRealisedPnLStocks(pnllist);
            var allpositionalTrades = pnllist.Where(x => x.Symbol == symbol).ToList();

            //Get Stock Current Details
            var stockcmplist = GoogleSheets.AllStocksSheetData("StockCMP");
            var stockdetails = stockcmplist.FirstOrDefault(x => x.Symbol == symbol);

            var dividendlist = GoogleSheets.GetDividendHistoryPnL("DividendHistory");
            var stockdividendlist = GoogleSheets.GetDividendHistoryStocks(dividendlist);
            var dividendDetails = dividendlist.Where(x => x.Symbol.Trim() == symbol).ToList();

            var dividendRawlist = GoogleSheets.GetDividendRawPnL("Dividend");
            var dividendRawDetails = dividendRawlist.Where(x => x.Symbol.Trim() == symbol).ToList();


            return View(new StockDetailsResponseModel()
            {
                stockTransactions = allStockTransactions,
                stockTrades = allpositionalTrades,
                stockdetails = stockdetails,
                dividendDetails = dividendDetails,
                dividendRawlist = dividendRawDetails
            });
        }
    }
}
