using System;
using System.Collections.Generic;

namespace GoogleSheetsAPI_Example
{
    class Program
    {
        static void Main(string[] args)
        {
           String spreadsheetId = "1DsTn1vCGNm38Kjcts6E6yBwd_IhmmMejjBXqkJIiCCU";
            //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "Sheets1!A2:C";
            GSheets sheets = new GSheets();
            IList<IList<Object>>  result = sheets.GetData(spreadsheetId, range);
            if (result != null && result.Count > 0)
            {
                Console.WriteLine("ID, Name, Income");
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
