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
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Example";

        static SheetsService GetService()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // Файл token.json хранит маркеры доступа и обновления для пользователя и создается автоматически при первом завершении потока авторизации.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Создаём сервис Google Sheets API
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        static void GetDataFromGoogleSheetByID(string spreadsheetId, string range)
        {
            SheetsService service = GetService();
            // Определим параметры запроса
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Печатаем ID, имя и счёт пользователя из таблицы
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("ID\tName\tAccount");
                foreach (var row in values)
                    Console.WriteLine($"{String.Format("{0,-10}", row[0])}\t{String.Format("{0,-8}", row[1])}\t{String.Format("{0,-8}", row[4])}");
            }
            else
                Console.WriteLine("Нет данных");
        }

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
            GetDataFromGoogleSheetByID("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", "Class Data!A2:E");
            //GetDataFromGoogleSheetByID("1d8RetomIySKDZp1e0bGSuDlKUEfiQQQR7uapifHABT0", "Class Data!A2:C");

            CreateNewSheet();
        }
    }
}
