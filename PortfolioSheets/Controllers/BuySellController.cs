using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class BuySellController : Controller
    {
        public IActionResult Index()
        {
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list = GoogleSheets.AllSheetData("Breakdown", out transactionList);
            list.Sort((x, y) => x.Name.CompareTo(y.Name));

            return View(list);
        }

        [HttpPost]
        public IActionResult Buy(Models.BuyModel buyModel)
        {
            if(ModelState.IsValid)
                RedirectToAction("Index");

            GoogleSheets.InsertBuyToSheet(buyModel);
            return RedirectToAction("Index");
        }
    }
}
