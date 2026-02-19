using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class PortfolioDairyController : Controller
    {
        public IActionResult Index()
        {
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list = GoogleSheets.AllBreakdownSheetData("Breakdown");
            list.Sort((x, y) => x.Date.CompareTo(y.Date));

            return View(list);
        }
    }
}
