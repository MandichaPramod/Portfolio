using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class FundsHistoryController : Controller
    {
        public IActionResult Index()
        {
            var list = GoogleSheets.GetFundsHistory("Funds");
            var stockslist = GoogleSheets.GetFundsHistoryAccounts(list);
            list.Sort((x, y) => x.Date.CompareTo(y.Date));

            var response = new FundsHistoryResponseModel();
            response.fundsHistory = list;
            response.fundsHistoryAccount = stockslist;
            return View(response);
        }
    }
}
