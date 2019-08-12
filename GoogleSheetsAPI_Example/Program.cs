using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetsAPI_Example
{
    class Program
    {

        static void CreateNewSheet()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = "1d8RetomIySKDZp1e0bGSuDlKUEfiQQQR7uapifHABT0";
            String range = "Sheet1!A10:D";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var x in values)
                {
                    //Rows.Add(x[0], x[1], x[2], x[3]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        static void Main(string[] args)
        {
           String spreadsheetId = "1DsTn1vCGNm38Kjcts6E6yBwd_IhmmMejjBXqkJIiCCU";
            //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "Sheets1!A2:C";
            GSheets sheets = new GSheets();
            IList<IList<Object>>  result = sheets.GetData(spreadsheetId, range);
            if (result != null && result.Count > 0)
            {
                Console.WriteLine("ID, Name, Major");
                foreach (var row in result)
                {
                    Console.WriteLine("{0}, {1}, {2}", row[0], row[1], row[2]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            Console.WriteLine("DataAdd:\n");
            sheets.AddData(spreadsheetId, range);
        }
    }
}
