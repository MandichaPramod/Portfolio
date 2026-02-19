using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class RealisedPnLController : Controller
    {
        public IActionResult Index()
        {
            var list = GoogleSheets.GetRealisedPnL("InActive");
            var stockslist = GoogleSheets.GetRealisedPnLStocks(list);

            var response = new RealisedPnLResponseModel();
            response.realisedPnLs = list;
            response.realisedPnLStocks = stockslist.Item1;
            response.DeliveryrealisedPnLStocks = stockslist.Item2;
            response.IntradayrealisedPnLStocks = stockslist.Item3;
            return View(response);
        }
    }
}
