using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetsAPI_Example
{
    class GSheets
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheets";
        public static SheetsService Connect()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //  Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public IList<IList<Object>> GetData(String spreadsheetId, String range)
        {
            SheetsService sheetsService = Connect();
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

            return values;
        }

        public void AddData(string spreadsheetId, string range, string[] data)
        {
            SheetsService sheetsService = Connect();

            ClearValuesRequest clearrequestBody = new ClearValuesRequest();
            //var result = sheetsService.Spreadsheets.Values.Clear(clearrequestBody, spreadsheetId, "Sheets1!A3:C").Execute();// Append( valueRange, spreadsheetId, "Sheets1!A6").Execute();

            ValueRange body = new ValueRange();
            List<Object> sub = new List<object>();
            sub.Add("qwerewtreytrjy");
            List<List<Object>> list1 = new List<List<object>>();
            list1.Add(sub);
            IList<IList<Object>> values = GetData(spreadsheetId, range);
            values.Clear();
            values.Add(data);
            body.Values = values;
            SpreadsheetsResource.ValuesResource.UpdateRequest request1 = sheetsService.Spreadsheets.Values.Update(body, spreadsheetId, range);
            request1.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            Console.WriteLine(request1.ValueInputOption);
            UpdateValuesResponse result = request1.Execute();

            Console.WriteLine(JsonConvert.SerializeObject(result));
        }


        public static Spreadsheet CreateNewSheet => Connect().Spreadsheets.Create(new Spreadsheet()).Execute();

        public void ShowData(object result)
        {
            if (result != null && ((IList<IList<object>>)result).Count > 0)
            {
                Console.WriteLine("ID, Name, Income");
                foreach (IList<object> row in (IList<IList<object>>)result)
                    Console.WriteLine("{0}, {1}, {2}", row[0], row[1], row[2]);
            }
            else
                Console.WriteLine("No data found.");
        }
    }
}