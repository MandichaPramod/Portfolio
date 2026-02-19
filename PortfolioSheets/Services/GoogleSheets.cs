using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.CodeAnalysis;
using Newtonsoft.Json;
using PortfolioSheets.Models;
using System.Diagnostics;

namespace PortfolioSheets.Services
{
    

    public static class GoogleSheets
    {
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static string ApplicationName = "MyApp";
        public static string SpreadsheetId = "1nqB6yvhMOPcxeBBFMrATWPthH4T9W5JheoIJR_O7wBY";
        public static SheetsService? service;
        public static List<BreakDownRow> breakdownstocks;
        public static List<RealisedPnLRow> realisedPnLStocks;
        public static List<DividendHistoryRow> dividendHistoryStocks;
        public static List<FundsHistoryRow> fundsHistoryStocks;
        public static List<MarketCapHistory> marketCapHistories;
        public static List<DividendRawRow> dividendRawStocks;
        public static List<StockCMPSheet> stocksCMP;

        private static SheetsService CreateConnection()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public static IList<IList<object>> GetDataFromSheet(string sheet,string sheetrange)
        {
            var service = CreateConnection();

            var range = $"{sheet}!{sheetrange}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;

            return values;
        }
        private static void GetBreakDownSheets(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:P");
            breakdownstocks = ParseBreakDownSheet(values);
        }

        public static List<RealisedPnLRow> GetRealisedPnL(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:Y");
            realisedPnLStocks = ParseRealisedPnL(values);
            return realisedPnLStocks;
        }

        public static List<DividendHistoryRow> GetDividendHistoryPnL(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:E");
            dividendHistoryStocks = ParseDividendHistory(values);
            return dividendHistoryStocks;
        }

        public static List<FundsHistoryRow> GetFundsHistory(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:E");
            fundsHistoryStocks = ParseFundsHistory(values);
            return fundsHistoryStocks;
        }

        public static List<MarketCapHistory> GetMarketCapHistory(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:S");
            marketCapHistories = ParseMarketCapHistory(values);
            return marketCapHistories;
        }

        public static List<DividendRawRow> GetDividendRawPnL(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A3:G");
            dividendRawStocks = ParseDividendRaw(values);
            return dividendRawStocks;
        }

        public static (List<RealisedPnLRow>,List<RealisedPnLRow>,List<RealisedPnLRow>) GetRealisedPnLStocks(List<RealisedPnLRow> realisedPnLRows)
        {
            var realisedPnLStocksWise = new List<RealisedPnLRow>();
            var DeliveryrealisedPnLStocksWise = new List<RealisedPnLRow>();
            var IntradayrealisedPnLStocksWise = new List<RealisedPnLRow>();
            foreach (var row in realisedPnLRows)
            {
                if (realisedPnLStocksWise.Any(x => x.Symbol == row.Symbol))
                {
                    var pf = realisedPnLStocksWise.First(x => x.Symbol == row.Symbol);
                    var quantity = pf.Quantity + row.Quantity;
                    var BuyPrice = ((pf.Quantity * pf.BuyPrice) + row.TotalCost) / quantity;
                    var Totalcost = BuyPrice * quantity;
                    var SellAvg = ((pf.Quantity * pf.SellPrice) + row.SellValue) / quantity;
                    var SellValue = SellAvg * quantity;

                    pf.Quantity = quantity;
                    pf.SellPrice = SellAvg;
                    pf.PnL = pf.PnL+row.PnL;
                    pf.Dividend = pf.Dividend+row.Dividend;
                    pf.BuyPrice = BuyPrice;
                    pf.TotalCost = Totalcost;
                    pf.SellValue = SellValue;
                    pf.Account = "";
                    pf.TotalCharges = pf.TotalCharges + row.TotalCharges;
                    pf.Brokerage = pf.Brokerage + row.Brokerage;
                    pf.STT = pf.STT + row.STT;
                    pf.TranCharges = pf.TranCharges + row.TranCharges;
                    pf.GST = pf.GST + row.GST;
                    pf.SEBICharges = pf.SEBICharges + row.SEBICharges;
                    pf.StampDuty = pf.StampDuty + row.StampDuty;
                    pf.OtherCharges = pf.OtherCharges + row.OtherCharges;
                    pf.PnLAfterCharges = pf.PnLAfterCharges + row.PnLAfterCharges;
                }
                else
                {
                    realisedPnLStocksWise.Add(new RealisedPnLRow()
                    {
                        Account = row.Account,
                        Symbol = row.Symbol,
                        Name = row.Name,
                        SellDate = row.SellDate,
                        Quantity = row.Quantity,
                        BuyPrice = row.BuyPrice,
                        SellPrice = row.SellPrice,
                        Dividend = row.Dividend,
                        TotalCost = row.TotalCost,
                        SellValue = row.SellValue,
                        PnL = row.PnL,
                        PnLAfterCharges = row.PnLAfterCharges,
                        TotalCharges = row.TotalCharges,
                        Brokerage = row.Brokerage,
                        STT = row.STT,
                        TranCharges = row.TranCharges,
                        GST = row.GST,
                        SEBICharges = row.SEBICharges,
                        StampDuty = row.StampDuty,
                        OtherCharges = row.OtherCharges
                    });
                }

                if(row.Type=="Delivery")
                {
                    if (DeliveryrealisedPnLStocksWise.Any(x => x.Symbol == row.Symbol))
                    {
                        var pf = DeliveryrealisedPnLStocksWise.First(x => x.Symbol == row.Symbol);
                        var quantity = pf.Quantity + row.Quantity;
                        var BuyPrice = ((pf.Quantity * pf.BuyPrice) + row.TotalCost) / quantity;
                        var Totalcost = BuyPrice * quantity;
                        var SellAvg = ((pf.Quantity * pf.SellPrice) + row.SellValue) / quantity;
                        var SellValue = SellAvg * quantity;

                        pf.Quantity = quantity;
                        pf.SellPrice = SellAvg;
                        pf.PnL = pf.PnL+row.PnL;
                        pf.Dividend = pf.Dividend+row.Dividend;
                        pf.BuyPrice = BuyPrice;
                        pf.TotalCost = Totalcost;
                        pf.SellValue = SellValue;
                        pf.Account = "";
                        pf.TotalCharges = pf.TotalCharges + row.TotalCharges;
                        pf.Brokerage = pf.Brokerage + row.Brokerage;
                        pf.STT = pf.STT + row.STT;
                        pf.TranCharges = pf.TranCharges + row.TranCharges;
                        pf.GST = pf.GST + row.GST;
                        pf.SEBICharges = pf.SEBICharges + row.SEBICharges;
                        pf.StampDuty = pf.StampDuty + row.StampDuty;
                        pf.OtherCharges = pf.OtherCharges + row.OtherCharges;
                        pf.PnLAfterCharges = pf.PnLAfterCharges + row.PnLAfterCharges;
                    }
                    else
                    {
                        DeliveryrealisedPnLStocksWise.Add(new RealisedPnLRow()
                        {
                            Account = row.Account,
                            Symbol = row.Symbol,
                            Name = row.Name,
                            SellDate = row.SellDate,
                            Quantity = row.Quantity,
                            BuyPrice = row.BuyPrice,
                            SellPrice = row.SellPrice,
                            Dividend = row.Dividend,
                            TotalCost = row.TotalCost,
                            SellValue = row.SellValue,
                            PnL = row.PnL,
                            PnLAfterCharges = row.PnLAfterCharges,
                            TotalCharges = row.TotalCharges,
                            Brokerage = row.Brokerage,
                            STT = row.STT,
                            TranCharges = row.TranCharges,
                            GST = row.GST,
                            SEBICharges = row.SEBICharges,
                            StampDuty = row.StampDuty,
                            OtherCharges = row.OtherCharges
                        });
                    }
                }
                else if(row.Type=="Intraday")
                {
                    if (IntradayrealisedPnLStocksWise.Any(x => x.Symbol == row.Symbol))
                    {
                        var pf = IntradayrealisedPnLStocksWise.First(x => x.Symbol == row.Symbol);
                        var quantity = pf.Quantity + row.Quantity;
                        var BuyPrice = ((pf.Quantity * pf.BuyPrice) + row.TotalCost) / quantity;
                        var Totalcost = BuyPrice * quantity;
                        var SellAvg = ((pf.Quantity * pf.SellPrice) + row.SellValue) / quantity;
                        var SellValue = SellAvg * quantity;

                        pf.Quantity = quantity;
                        pf.SellPrice = SellAvg;
                        pf.PnL = pf.PnL+row.PnL;
                        pf.Dividend = pf.Dividend+row.Dividend;
                        pf.BuyPrice = BuyPrice;
                        pf.TotalCost = Totalcost;
                        pf.SellValue = SellValue;
                        pf.Account = "";
                        pf.TotalCharges = pf.TotalCharges + row.TotalCharges;
                        pf.Brokerage = pf.Brokerage + row.Brokerage;
                        pf.STT = pf.STT + row.STT;
                        pf.TranCharges = pf.TranCharges + row.TranCharges;
                        pf.GST = pf.GST + row.GST;
                        pf.SEBICharges = pf.SEBICharges + row.SEBICharges;
                        pf.StampDuty = pf.StampDuty + row.StampDuty;
                        pf.OtherCharges = pf.OtherCharges + row.OtherCharges;
                        pf.PnLAfterCharges = pf.PnLAfterCharges + row.PnLAfterCharges;
                    }
                    else
                    {
                        IntradayrealisedPnLStocksWise.Add(new RealisedPnLRow()
                        {
                            Account = row.Account,
                            Symbol = row.Symbol,
                            Name = row.Name,
                            SellDate = row.SellDate,
                            Quantity = row.Quantity,
                            BuyPrice = row.BuyPrice,
                            SellPrice = row.SellPrice,
                            Dividend = row.Dividend,
                            TotalCost = row.TotalCost,
                            SellValue = row.SellValue,
                            PnL = row.PnL,
                            PnLAfterCharges = row.PnLAfterCharges,
                            TotalCharges = row.TotalCharges,
                            Brokerage = row.Brokerage,
                            STT = row.STT,
                            TranCharges = row.TranCharges,
                            GST = row.GST,
                            SEBICharges = row.SEBICharges,
                            StampDuty = row.StampDuty,
                            OtherCharges = row.OtherCharges
                        });
                    }
                }
            }
            return (realisedPnLStocksWise,DeliveryrealisedPnLStocksWise,IntradayrealisedPnLStocksWise);
        }

        public static List<DividendHistoryRow> GetDividendHistoryStocks(List<DividendHistoryRow> dividendHistoryRows)
        {
            var dividendHistoryStocksWise = new List<DividendHistoryRow>();
            foreach (var row in dividendHistoryRows)
            {
                if (dividendHistoryStocksWise.Any(x => x.Symbol == row.Symbol))
                {
                    var pf = dividendHistoryStocksWise.First(x => x.Symbol == row.Symbol);
                   
                    pf.Dividend = pf.Dividend + row.Dividend;
                }
                else
                {
                    dividendHistoryStocksWise.Add(new DividendHistoryRow()
                    {
                        Symbol = row.Symbol,
                        Company = row.Company,
                        Date = row.Date,
                        Dividend = row.Dividend,
                        Remarks = row.Remarks
                    });
                }
            }
            return dividendHistoryStocksWise;
        }

        public static List<FundsHistoryRow> GetFundsHistoryAccounts(List<FundsHistoryRow> fundsHistoryRows)
        {
            var fundsHistoryAccountWise = new List<FundsHistoryRow>();
            foreach (var row in fundsHistoryRows)
            {
                if (fundsHistoryAccountWise.Any(x => x.Account == row.Account))
                {
                    var pf = fundsHistoryAccountWise.First(x => x.Account == row.Account);

                    pf.Debit = pf.Debit + row.Debit;
                    pf.Credit = pf.Credit + row.Credit;
                }
                else
                {
                    fundsHistoryAccountWise.Add(new FundsHistoryRow()
                    {
                        Account = row.Account,
                        Date = row.Date,
                        Debit = row.Debit,
                        Credit = row.Credit,
                        Remarks = row.Remarks
                    });
                }
            }
            return fundsHistoryAccountWise;
        }

        public static List<Portfolio> AllSheetData(string sheet,out Dictionary<string, List<BreakDownRow>> stockTransactions)
        {
            List<Portfolio> portfolios = new List<Portfolio>();
            stockTransactions = new Dictionary<string, List<BreakDownRow>>();

            //if (breakdownstocks == null)
                GetBreakDownSheets(sheet);
            //GetCMP();

            foreach (var row in breakdownstocks)
            {
                if (portfolios.Any(x => x.Symbol == row.Symbol))
                {
                    double BuyPrice = 0;
                    var pf=portfolios.First(x => x.Symbol == row.Symbol);
                    var quantity = pf.Quantity + row.Quantity;
                    BuyPrice = ((pf.Quantity * pf.BuyPrice) + row.TotalCost) / quantity;
                    var Totalcost = BuyPrice * quantity;
                    var CurrentValue = pf.CurrentPrice * quantity;
                    var Dividend = pf.Dividend + row.Dividend - row.DividendReinvestment;
                    var PnL = CurrentValue - Totalcost + Dividend;

                    pf.Quantity = quantity;
                    pf.CurrentValue = CurrentValue;
                    pf.PnL = PnL;
                    pf.Dividend = Dividend;
                    pf.BuyPrice = BuyPrice;
                    pf.TotalCost = Totalcost;
                    pf.Change = pf.Change+ row.Change;

                    var transactionlist=stockTransactions[row.Symbol];
                    transactionlist.Add(row);
                    stockTransactions[row.Symbol] = transactionlist;
                }
                else
                {
                    double BuyPrice = row.TotalCost / row.Quantity;
                    portfolios.Add(new Portfolio
                    {
                        Symbol = row.Symbol,
                        Name = row.Name,
                        Quantity = row.Quantity,
                        BuyPrice = BuyPrice,
                        TotalCost = row.TotalCost,
                        CurrentValue = row.CurrentValue,
                        CurrentPrice = row.CurrentValue/row.Quantity,
                        Dividend = row.Dividend-row.DividendReinvestment,
                        PnL = row.CurrentValue-row.TotalCost+(row.Dividend - row.DividendReinvestment),
                        Change = row.Change,
                    });

                    stockTransactions[row.Symbol] = new List<BreakDownRow>() { row };
                }

            }

            return portfolios;
        }

        public static List<BreakDownRow> AllBreakdownSheetData(string sheet)
        {
            GetBreakDownSheets(sheet);
            //GetCMP();

            return breakdownstocks;
        }

        public static List<StockCMPSheet> AllStocksSheetData(string sheet)
        {
            var stockcmp=GetStockCMPSheets(sheet);
            //GetCMP();

            return stockcmp;
        }

        private static List<StockCMPSheet> GetStockCMPSheets(string sheet)
        {
            var values = GetDataFromSheet(sheet, "A2:R");
            var stockcmp = ParseStockCMPSheet(values);
            return stockcmp;
        }

        private static List<StockCMPSheet> ParseStockCMPSheet(IList<IList<object>> values)
        {
            var stockcmprows = new List<StockCMPSheet>();
            string symbol = "";
            try
            {

                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;
                    var stockcmprow = new StockCMPSheet();
                    stockcmprow.Symbol = row[0].ToString();
                    stockcmprow.Name = row[1].ToString();
                    stockcmprow.CMP = row[2].ToString().Contains("#N/A")? 0.0 : Double.Parse(row[2].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.OldCMP = row[3].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[3].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Open = row[4].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[4].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.High = row[5].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[5].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Low = row[6].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[6].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Volume = row[7].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[7].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.MarketCap = row[8].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[8].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.PriceToEarnings = row[9].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[9].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.EarningsPerShare = row[10].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[10].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.High52W = row[11].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[11].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Low52W = row[12].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[12].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Change = row[13].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[13].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.ChangePercent = row[14].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[14].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.CloseYest = row[15].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[15].ToString().Replace("₹", "").Replace("$", ""));
                    stockcmprow.Shares = row[16].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[16].ToString().Replace("₹", "").Replace("$", ""));
                    //stockcmprow.LastTradedTime = DateTime.Parse(row[17].ToString());

                    stockcmprows.Add(stockcmprow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return stockcmprows;
        }

        public static List<Portfolio> GetSheetsData(string sheet,string PortfolioName)
        {
            List<Portfolio> portfolios = new List<Portfolio>();

            if (breakdownstocks == null)
                GetBreakDownSheets(sheet);

            foreach (var row in breakdownstocks)
            {
                if (row.Portfolio != PortfolioName)
                    continue;
                if (portfolios.Any(x => x.Symbol == row.Symbol && x.Account == row.Account))
                {
                    double BuyPrice = 0;
                    var pf = portfolios.First(x => x.Symbol == row.Symbol && x.Account == row.Account);
                    var quantity = pf.Quantity + row.Quantity;
                    BuyPrice = ((pf.Quantity * pf.BuyPrice) + row.TotalCost) / quantity;
                    var Totalcost = BuyPrice * quantity;
                    var CurrentValue = pf.CurrentPrice * quantity;
                    var Dividend = pf.Dividend + row.Dividend - row.DividendReinvestment;
                    var PnL = CurrentValue - Totalcost + Dividend;

                    pf.Quantity = quantity;
                    pf.CurrentValue = CurrentValue;
                    pf.PnL = PnL;
                    pf.Dividend = Dividend;
                    pf.BuyPrice = BuyPrice;
                    pf.TotalCost = Totalcost;
                    pf.Change = pf.Change + row.Change;
                }
                else
                {
                    double BuyPrice = row.TotalCost / row.Quantity;
                    portfolios.Add(new Portfolio
                    {
                        Account = row.Account,
                        Symbol = row.Symbol,
                        Name = row.Name,
                        Quantity = row.Quantity,
                        BuyPrice = BuyPrice,
                        TotalCost = row.TotalCost,
                        CurrentValue = row.CurrentValue,
                        CurrentPrice = row.CurrentValue / row.Quantity,
                        Dividend = row.Dividend - row.DividendReinvestment,
                        PnL = row.CurrentValue - row.TotalCost + (row.Dividend - row.DividendReinvestment),
                        Change = row.Change,
                    });
                }

            }

            return portfolios;
        }


        private static List<BreakDownRow> ParseBreakDownSheet(IList<IList<object>> values)
        {
            var breakdownrows = new List<BreakDownRow>();
            string symbol="";
            try
            {

                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;
                    if (Double.Parse(row[4].ToString()) == 0)
                        continue;
                    symbol= row[0].ToString(); ;
                    var breakdownrow = new BreakDownRow();
                    breakdownrow.Symbol = row[0].ToString();
                    breakdownrow.Name = row[1].ToString();
                    breakdownrow.Account = row[2].ToString();
                    breakdownrow.Date = DateTime.Parse(row[3].ToString());
                    breakdownrow.Quantity = Double.Parse(row[4].ToString());
                    breakdownrow.BuyPrice = Double.Parse(row[5].ToString().Replace("₹", ""));
                    breakdownrow.TotalCost = Double.Parse(row[6].ToString().Replace("₹", ""));
                    breakdownrow.CurrentValue = Double.Parse(row[7].ToString().Replace("₹", ""));
                    breakdownrow.Dividend = Double.Parse(row[9].ToString());
                    breakdownrow.DividendReinvestment = (row[10].ToString() == "") ? 0 : Double.Parse(row[10].ToString().Replace("₹", ""));
                    breakdownrow.Portfolio = row[11].ToString();
                    breakdownrow.Tag = row[12].ToString();
                    breakdownrow.Ageing = Int32.Parse(row[13].ToString());
                    breakdownrow.Change = Double.Parse(row[15].ToString().Replace("₹", ""));
                    breakdownrows.Add(breakdownrow);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return breakdownrows;
        }

        private static List<RealisedPnLRow> ParseRealisedPnL(IList<IList<object>> values)
        {
            var pnlrows = new List<RealisedPnLRow>();
            string symbol = "";
            try
            {

                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;
                    if (row[1].ToString() == "NVDA")
                        Console.WriteLine("");
                    var pnlrow = new RealisedPnLRow();
                    pnlrow.Account = row[0].ToString();
                    pnlrow.Symbol = row[1].ToString();
                    pnlrow.Name = row[2].ToString();
                    pnlrow.SellDate = DateTime.Parse(row[3].ToString());
                    pnlrow.Quantity = Double.Parse(row[4].ToString());
                    pnlrow.BuyPrice = Double.Parse(row[5].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.SellPrice = Double.Parse(row[6].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.Dividend = (row[7].ToString() == "") ? 0 : Double.Parse(row[7].ToString().Replace("₹", ""));
                    pnlrow.TotalCost = Double.Parse(row[8].ToString().Replace("₹", ""));
                    pnlrow.SellValue = Double.Parse(row[9].ToString().Replace("₹", ""));
                    pnlrow.PnL = Double.Parse(row[10].ToString().Replace("₹", ""));
                    pnlrow.PnLAfterCharges = Double.Parse(row[12].ToString().Replace("₹", ""));
                    pnlrow.TotalCharges = Double.Parse(row[14].ToString().Replace("₹", ""));
                    pnlrow.Brokerage = Double.Parse(row[16].ToString().Replace("₹", ""));
                    pnlrow.STT = Double.Parse(row[17].ToString().Replace("₹", ""));
                    pnlrow.TranCharges = Double.Parse(row[18].ToString().Replace("₹", ""));
                    pnlrow.GST = Double.Parse(row[19].ToString().Replace("₹", ""))+ Double.Parse(row[20].ToString().Replace("₹", ""));
                    pnlrow.SEBICharges = Double.Parse(row[21].ToString().Replace("₹", ""));
                    pnlrow.StampDuty = Double.Parse(row[22].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.OtherCharges = Double.Parse(row[23].ToString().Replace("₹", ""));
                    pnlrow.Type = row[24].ToString();

                    pnlrows.Add(pnlrow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return pnlrows;
        }

        private static List<DividendHistoryRow> ParseDividendHistory(IList<IList<object>> values)
        {
            var rows = new List<DividendHistoryRow>();
            string symbol = "";
            try
            {
                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;

                    var pnlrow = new DividendHistoryRow();
                    pnlrow.Symbol = row[0].ToString();
                    pnlrow.Company = row[1].ToString();
                    pnlrow.Date = DateTime.Parse(row[2].ToString());
                    pnlrow.Dividend = Double.Parse(row[3].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.Remarks = row[4].ToString();

                    rows.Add(pnlrow);

                   
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return rows;
        }

        private static List<FundsHistoryRow> ParseFundsHistory(IList<IList<object>> values)
        {
            var rows = new List<FundsHistoryRow>();
            string symbol = "";
            try
            {
                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;

                    var pnlrow = new FundsHistoryRow();
                    pnlrow.Account = row[0].ToString();
                    pnlrow.Date = DateTime.Parse(row[1].ToString());
                    pnlrow.Debit = Double.Parse(row[2].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.Credit = Double.Parse(row[3].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.Remarks = row[4].ToString();

                    rows.Add(pnlrow);


                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return rows;
        }

        private static List<MarketCapHistory> ParseMarketCapHistory(IList<IList<object>> values)
        {
            var rows = new List<MarketCapHistory>();
            string symbol = "";
            try
            {
                for(int i=0;i<values.Count;i++)
                {
                    var row = new MarketCapHistory();
                    if (values[i].Count == 0)
                        continue;

                    row.Month = DateTime.Parse(values[i][0].ToString());

                    var investedRow = new MarketCapHistoryRow();
                    investedRow.StockMarket = Double.Parse(values[i][2].ToString().Replace("₹", "").Replace("$", ""));

                    ++i;
                    var currentRow = new MarketCapHistoryRow();
                    currentRow.StockMarket = Double.Parse(values[i][2].ToString().Replace("₹", "").Replace("$", ""));

                    row.Invested = investedRow;
                    row.Current = currentRow;
                    row.StockMarketPercent = System.Math.Round(((currentRow.StockMarket - investedRow.StockMarket) * 100 / investedRow.StockMarket), 2);
                    rows.Add(row);


                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return rows;
        }
        private static List<DividendRawRow> ParseDividendRaw(IList<IList<object>> values)
        {
            var rows = new List<DividendRawRow>();
            string symbol = "";
            try
            {
                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;

                    var pnlrow = new DividendRawRow();
                    pnlrow.Symbol = row[0].ToString();
                    pnlrow.Date = DateTime.Parse(row[1].ToString());
                    pnlrow.Quantity = Double.Parse(row[2].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.DividendPerShare = Double.Parse(row[3].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.TotalDividend = Double.Parse(row[4].ToString().Replace("₹", "").Replace("$", ""));
                    pnlrow.StockPrize = row[5].ToString().Contains("#N/A") ? 0.0 : Double.Parse(row[5].ToString().Replace("₹", "").Replace("$", ""));

                    rows.Add(pnlrow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return rows;
        }

        private static void GetCMP()
        {
            var service = CreateConnection();
            stocksCMP = new List<StockCMPSheet>();

            string sheet = "StockCMP";
            var range = $"{sheet}!A2:D";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            int idx = 0;

            foreach(var row in values)
            {
                var stockCMP = new StockCMPSheet();
                stockCMP.Symbol = row[0].ToString();
                if (row[1].ToString()=="#N/A")
                {
                    stockCMP.CMP = Double.Parse(row[2].ToString().Replace("₹", ""));
                }
                else
                {
                    stockCMP.CMP = Double.Parse(row[1].ToString().Replace("₹", ""));
                    UpdateSheet(service, sheet, "C" + (idx + 2), stockCMP.CMP.ToString());
                    UpdateSheet(service, sheet, "D" + (idx + 2), DateTime.Now.ToString());
                }
                stockCMP.OldCMP = Double.Parse(row[2].ToString().Replace("₹", ""));
                //stockCMP.Lastupdate = DateTime.Parse(row[3].ToString());

                stocksCMP.Add(stockCMP);
                idx++;
            }
        }
        public static void UpdateSheet(SheetsService service,string sheet,string cell,string content)
        {
            var range = $"{sheet}!{cell}";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { content };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
        }
        public static void InsertBuyToSheet(BuyModel buyModel)
        {
            var sheet = "Breakdown";
            var service = CreateConnection();

            var range = $"{sheet}!A:P";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            var index = values.Count+1;

            range = $"{sheet}!A{index}:N{index}";
            var valueRange = new ValueRange();

            if (buyModel.Portfolio == "US Stocks" || buyModel.Portfolio == "US Free Stocks")
            {
                var oblist = new List<object>() { buyModel.Symbol, $"=GOOGLEFINANCE(A{index},\"NAME\")", buyModel.Account, "=Date(" + buyModel.BuyDate.Year + "," + buyModel.BuyDate.Month + "," + buyModel.BuyDate.Day + ")", buyModel.Qty, buyModel.BuyPrice, $"=E{index} * F{index} * GOOGLEFINANCE(\"CURRENCY:USDINR\")", $"=E{index}*GOOGLEFINANCE(A{index},\"PRICE\")*GOOGLEFINANCE(\"CURRENCY:USDINR\")", $"=H{index}-G{index}", $"=(SUMIFS(Dividend!D:D, Dividend!A:A, A{index}, Dividend!B:B,\">=\"&D{index})*E{index})", "", buyModel.Portfolio, "", $"=TODAY()-D{index}", "", $"=E{index} * GOOGLEFINANCE(A{index},\"CHANGE\")*GOOGLEFINANCE(\"CURRENCY:USDINR\")" };
                valueRange.Values = new List<IList<object>> { oblist };
            }
            else
            {
                var oblist = new List<object>() { buyModel.Symbol, $"=GOOGLEFINANCE(A{index},\"NAME\")", buyModel.Account, "=Date(" + buyModel.BuyDate.Year + "," + buyModel.BuyDate.Month + "," + buyModel.BuyDate.Day + ")", buyModel.Qty, buyModel.BuyPrice, $"=E{index} * F{index}", $"=E{index}*GOOGLEFINANCE(A{index},\"PRICE\")", $"=H{index}-G{index}", $"=(SUMIFS(Dividend!D:D, Dividend!A:A, A{index}, Dividend!B:B,\">=\"&D{index})*E{index})", "", buyModel.Portfolio, "", $"=TODAY()-D{index}", "", $"=E{index} * GOOGLEFINANCE(A{index},\"CHANGE\")" };
                valueRange.Values = new List<IList<object>> { oblist };
            }

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }
    }
}
