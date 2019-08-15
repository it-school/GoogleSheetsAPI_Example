using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;

namespace GoogleSheetsAPI_Example
{
    class Program
    {
        static void Main(string[] args)
        {
            string spreadsheetId = "1DsTn1vCGNm38Kjcts6E6yBwd_IhmmMejjBXqkJIiCCU";
            string range = "Sheets1!A2:C";
            GSheets sheets = new GSheets();
            Random r = new Random();

            IList<IList<Object>>  result = sheets.GetData(spreadsheetId, range);
            sheets.ShowData(result);

            Console.WriteLine("DataAdd:");
            string[] data = { (result.Count+1).ToString(), "NewData"+ (result.Count + 1), r.Next().GetHashCode().ToString()};
            sheets.AddData(spreadsheetId, "Sheets1!A"+(result.Count+1), data);

            sheets.ShowData(sheets.GetData(spreadsheetId, range));

            Console.WriteLine("Создан новый файл: " + GSheets.CreateNewSheet.SpreadsheetUrl);
        }
    }
}
