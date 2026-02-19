using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;

namespace PortfolioSheets.Controllers
{
    public class IVDairyController : Controller
    {
        public IActionResult Index(bool bonusSplitAdjusted = false)
        {
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list = GoogleSheets.AllBreakdownSheetData("Breakdown");
            list.Sort((x, y) => x.Date.CompareTo(y.Date));

            if (bonusSplitAdjusted)
            {
                // Find all bonus/split items and adjust previous buy values
                var bonusSplitItems = list.Where(x => x.Tag == "Bonus" || x.Tag == "Split").ToList();

                foreach (var bonusItem in bonusSplitItems)
                {
                    // Find all previous items with same Symbol, Portfolio, and Account
                    var previousItems = list.Where(x =>
                        x.Symbol == bonusItem.Symbol &&
                        x.Portfolio == bonusItem.Portfolio &&
                        x.Account == bonusItem.Account &&
                        x.Date < bonusItem.Date &&
                        (x.Tag != "Bonus" && x.Tag != "Split")
                    ).ToList();

                    // Adjust the buy price and total cost of previous items
                    if (previousItems.Count > 0)
                    {
                        double adjustmentFactor = 1;

                        if (bonusItem.Tag == "Bonus")
                        {
                            // For bonus: new quantity = old quantity + bonus quantity
                            // Adjusted buy price = old total cost / new quantity
                            double totalQuantityAfterBonus = previousItems.Sum(x => x.Quantity) + bonusItem.Quantity;
                            double totalCostBeforeBonus = previousItems.Sum(x => x.TotalCost);
                            
                            foreach (var prevItem in previousItems)
                            {
                                double quantity = prevItem.Quantity;
                                double newBuyPrice = totalCostBeforeBonus / totalQuantityAfterBonus;
                                prevItem.BuyPrice = newBuyPrice;
                                prevItem.Quantity = prevItem.Quantity + (bonusItem.Quantity * prevItem.Quantity / previousItems.Sum(x => x.Quantity));
                                prevItem.TotalCost = prevItem.Quantity * newBuyPrice;
                                prevItem.CurrentValue = (prevItem.CurrentValue/quantity) * prevItem.Quantity;
                            }
                        }
                        else if (bonusItem.Tag == "Split")
                        {
                            // For split: adjustment factor = old quantity / new quantity
                            // New buy price = old buy price * adjustment factor
                            /*adjustmentFactor = bonusItem.Quantity / previousItems.Sum(x => x.Quantity);

                            foreach (var prevItem in previousItems)
                            {                       
                                prevItem.Quantity =  prevItem.Quantity + (prevItem.Quantity * adjustmentFactor;
                                prevItem.BuyPrice = prevItem.BuyPrice * adjustmentFactor;
                                prevItem.TotalCost = Quantity * prevItem.BuyPrice;
                            }*/
                            double totalQuantityAfterBonus = previousItems.Sum(x => x.Quantity) + bonusItem.Quantity;
                            double totalCostBeforeBonus = previousItems.Sum(x => x.TotalCost);
                            
                            foreach (var prevItem in previousItems)
                            {
                                double quantity = prevItem.Quantity;
                                double newBuyPrice = totalCostBeforeBonus / totalQuantityAfterBonus;
                                prevItem.BuyPrice = newBuyPrice;
                                prevItem.Quantity = prevItem.Quantity + (bonusItem.Quantity * prevItem.Quantity / previousItems.Sum(x => x.Quantity));
                                prevItem.TotalCost = prevItem.Quantity * newBuyPrice;
                                prevItem.CurrentValue = (prevItem.CurrentValue/quantity) * prevItem.Quantity;
                            }
                        }
                    }
                }

                // Remove bonus/split items from the list
                list.RemoveAll(x => x.Tag == "Bonus" || x.Tag == "Split");

                // Remove all Dividend Reinvestment items
                list.RemoveAll(x => x.Tag == "Dividend Reinvestment");
            }
            ViewData["BonusSplitAdjusted"] = bonusSplitAdjusted;

            return View(list);
        }
    }
}
