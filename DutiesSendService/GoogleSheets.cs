using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace DutiesSendService
{
    public class GoogleSheets
    {
        private UserCredential credential;

        private string[] scopes = { SheetsService.Scope.Spreadsheets };

        public AccessAccountConfiguration AccessAccount { get; set; }

        public List<Duty> GetNewDuties(string spreadSheetId, string sheetName)
        {
            Authorize();

            var service = GetGoogleSheetsService();
            
            var request =
                service.Spreadsheets.Values.Get(spreadSheetId, sheetName);

            var response = request.Execute();

            var newDuties = response.Values.Where(x => x.Count >= 21 && x.ElementAt(21).ToString().ToLower() == "выдан")
                .Select(x => new Duty
                {
                    RowNumber = response.Values.IndexOf(x) + 1,
                    CustomerName = x.ElementAtOrDefault(0)?.ToString(),
                    PhoneLine = x.ElementAtOrDefault(1)?.ToString(),
                    TaskId = x.ElementAtOrDefault(2)?.ToString(),
                    SubscriberLineType = x.ElementAtOrDefault(3)?.ToString(),
                    Account = x.ElementAtOrDefault(4)?.ToString(),
                    Service = x.ElementAtOrDefault(5)?.ToString(),
                    ADSLType = x.ElementAtOrDefault(6)?.ToString(),
                    ADSLPort = x.ElementAtOrDefault(7)?.ToString(),
                    B6Network = x.ElementAtOrDefault(9)?.ToString(),
                    Region = x.ElementAtOrDefault(10)?.ToString(),
                    City = x.ElementAtOrDefault(11)?.ToString(),
                    Term = x.ElementAtOrDefault(12)?.ToString(),
                    CustomerAddress = x.ElementAtOrDefault(13)?.ToString(),
                    ContactName = x.ElementAtOrDefault(14)?.ToString(),
                    ContactPhone = x.ElementAtOrDefault(15)?.ToString(),
                    DutyCreator = x.ElementAtOrDefault(16)?.ToString(),
                    CreationDate = x.ElementAtOrDefault(17)?.ToString(),
                    ExpirationDate = x.ElementAtOrDefault(18)?.ToString(),
                    LinearDataForOTA = x.ElementAtOrDefault(19)?.ToString(),
                    NotesFromUkt = x.ElementAtOrDefault(20)?.ToString(),
                    Status = x.ElementAtOrDefault(21)?.ToString()
                }).ToList();

            return newDuties;
        }

        public void UpdateDutyStatus(string spreadSheetId, string sheetName, string cellName)
        {
            Authorize();

            var service = GetGoogleSheetsService();

            
            var cell = $"{sheetName}!{cellName}";
            var valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";

            var oblist = new List<object> { "В работе" };
            valueRange.Values = new List<IList<object>> { oblist };

            var update = service.Spreadsheets.Values.Update(valueRange, spreadSheetId, cell);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var response = update.Execute();
        }

        private void Authorize()
        {
            using (var stream = new FileStream($"{AccessAccount.authTokenFolder}\\client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                    scopes, "user", CancellationToken.None,
                    new FileDataStore($"{AccessAccount.authTokenFolder}\\sheets.googleapis.com-duties-service.json",
                        true)).Result;
            }
        }

        private SheetsService GetGoogleSheetsService()
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = AccessAccount.applicationName
            });
            return service;
        }
    }
}
