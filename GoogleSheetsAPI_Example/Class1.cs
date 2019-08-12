using Google.Apis.Sheets.v4;
using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace GoogleSheets
{
    class GSheets
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheets";
        SheetsService Connect()
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

        public void AddData(String spreadsheetId, String range)
        {
            SheetsService sheetsService = Connect();
            //sheetsService.Spreadsheets.Create(new Spreadsheet()).Execute(); ;
            
            ClearValuesRequest clearrequestBody = new ClearValuesRequest();
            //var result = sheetsService.Spreadsheets.Values.Clear(clearrequestBody, spreadsheetId, "Sheets1!A3:C").Execute();// Append( valueRange, spreadsheetId, "Sheets1!A6").Execute();



            string range1 = "Sheets1!A3";  // TODO: Update placeholder value.

            ValueRange body = new ValueRange();
            List<Object> sub = new List<object>();
            sub.Add("qwerewtreytrjy");
            List<List<Object>> list1 = new List<List<object>>();
            list1.Add(sub);
            IList<IList<Object>> values = GetData(spreadsheetId, range);
            values.Clear();
            values.Add(new List<Object>() { "dhxfhgkf", "Sgdzhtjydufky"});
            body.Values = values;
            SpreadsheetsResource.ValuesResource.UpdateRequest request1 = sheetsService.Spreadsheets.Values.Update(body, spreadsheetId, range1);
            request1.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            Console.WriteLine(request1.ValueInputOption);
            UpdateValuesResponse result = request1.Execute();

            Console.WriteLine(JsonConvert.SerializeObject(result));



        }

    }
}
