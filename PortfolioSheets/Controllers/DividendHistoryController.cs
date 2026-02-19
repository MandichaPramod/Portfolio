using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class DividendHistoryController : Controller
    {
        public IActionResult Index()
        {
            var list = GoogleSheets.GetDividendHistoryPnL("DividendHistory");
            var stockslist = GoogleSheets.GetDividendHistoryStocks(list);
            list.Sort((x, y) => x.Date.CompareTo(y.Date));

            var response = new DividendHistoryResponseModel();
            response.dividendHistory = list;
            response.dividendHistoryStocks = stockslist;
            return View(response);
        }
    }
}
