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
    

    public static class ExternalSheet
    {
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static string ApplicationName = "MyApp";
        public static SheetsService? service;
        public static List<ExternalCorporateEvent> events;

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

        public static IList<IList<object>> GetDataFromSheet(string sheet,string sheetrange,string spreadSheetId)
        {
            var service = CreateConnection();

            var range = $"{sheet}!{sheetrange}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadSheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;

            return values;
        }
        public static List<ExternalCorporateEvent> GetCorporateEvents(string sheet)
        {

            var values = GetDataFromSheet(sheet, "A2:J", "1yTedLYGJ8z-X2L7pNPyRnzJpelzzuKH_26vZSBfFjIY");
            return ParseCorporateEvents(values);
        }

        private static List<ExternalCorporateEvent> ParseCorporateEvents(IList<IList<object>> values)
        {
            var eventrows = new List<ExternalCorporateEvent>();
            string symbol = "";
            try
            {

                foreach (var row in values)
                {
                    if (row.Count == 0)
                        continue;

                    symbol = row[0].ToString(); ;
                    var eventRow = new ExternalCorporateEvent();
                    eventRow.Symbol = "NSE:"+row[1].ToString();
                    eventRow.Name = row[0].ToString();
                    eventRow.Segment = row[2].ToString();
                    eventRow.Type = row[3].ToString();
                    eventRow.Remarks = row[4].ToString();
                    eventRow.Status = row[5].ToString();
                    eventRow.BeforeDate = DateTime.Parse(row[6].ToString());
                    eventRow.ExDate = DateTime.Parse(row[7].ToString());
                    eventRow.Ratio = row[8].ToString();
                    eventRow.Cost = row[9].ToString();

                    eventrows.Add(eventRow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(symbol);
            }
            return eventrows;
        }

        internal static void UpdateEventInfo(List<ExternalCorporateEvent> events, List<Portfolio> transactionList)
        {
            foreach(ExternalCorporateEvent e in events)
            {
                if(transactionList.Any(x=>x.Symbol == e.Symbol))
                {
                    var list = transactionList.Where(x=>x.Symbol == e.Symbol).ToList();
                    if (e.ExDate >= DateTime.Today.AddDays(-15) && e.ExDate <= DateTime.Today.AddDays(15))
                    {
                        foreach(var item in list)
                        {
                            item.Events = new List<CorpEvent>
                            {
                                new CorpEvent(){ Type = e.Type, Value = e.ExDate.ToShortDateString()}
                            };
                        }
                    }
                }
            }
        }
    }
}
