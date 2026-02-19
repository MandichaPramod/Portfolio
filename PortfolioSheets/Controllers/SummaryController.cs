using Microsoft.AspNetCore.Mvc;
using PortfolioSheets.Models;
using PortfolioSheets.Services;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

namespace PortfolioSheets.Controllers
{
    public class SummaryController : Controller
    {
        private readonly ILogger<SummaryController> _logger;

        public SummaryController(ILogger<SummaryController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var transactionList = new Dictionary<string, List<BreakDownRow>>();
            var list=GoogleSheets.AllSheetData("Breakdown", out transactionList);

            //var events = ExternalSheet.GetCorporateEvents("FY 2025-26");
            //ExternalSheet.UpdateEventInfo(events, list);


            var allLists = new Dictionary<string, List<Portfolio> > { };
            list.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "Consolidated"]= list;

            var freestockslist = GoogleSheets.GetSheetsData("Breakdown", "Free Stock");
            freestockslist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists["Free Stock"]=freestockslist ;

            var longtermlist = GoogleSheets.GetSheetsData("Breakdown", "Long Term");
            longtermlist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "Long Term"]=longtermlist ;

            var mutualfundslist = GoogleSheets.GetSheetsData("Breakdown", "Mutual Fund");
            mutualfundslist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "Mutual Fund"] =mutualfundslist ;

            var individualStocklist = GoogleSheets.GetSheetsData("Breakdown", "Stock SIP");
            individualStocklist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "Stock SIP"] =individualStocklist ;

            var indexSIPlist = GoogleSheets.GetSheetsData("Breakdown", "ETF SIP");
            indexSIPlist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "ETF SIP"] =indexSIPlist ;

            var tradinglist = GoogleSheets.GetSheetsData("Breakdown", "Trading");
            tradinglist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "Trading"] =tradinglist ;

            var USFreelist = GoogleSheets.GetSheetsData("Breakdown", "US Free Stocks");
            USFreelist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "US Free Stocks"] =USFreelist ;

            var USStockslist = GoogleSheets.GetSheetsData("Breakdown", "US Stocks");
            USStockslist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists[ "US Stocks"] =USStockslist ;

            var USETFlist = GoogleSheets.GetSheetsData("Breakdown", "US ETF");
            USETFlist.Sort((x, y) => x.Name.CompareTo(y.Name));
            allLists["US ETF"] = USETFlist;

            var marketCapHistory = GoogleSheets.GetMarketCapHistory("Networth OverView");

            return View(new PortfolioResponseModel()
            {
                stockTransactions = transactionList,
                AllLists = allLists,
                MarketCapHistory = marketCapHistory
            });
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}