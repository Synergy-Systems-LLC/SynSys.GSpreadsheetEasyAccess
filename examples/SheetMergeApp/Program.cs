using Newtonsoft.Json;
using SynSys.GSpreadsheetEasyAccess;
using System;
using System.Collections.Generic;
using System.IO;

namespace SheetMergeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Запуск {nameof(SheetMergeApp)}\n");

            var connector = new Connector() { ApplicationName = "Get google sheet data", };
 
            bool isConnectionSuccessful = connector.TryToCreateConnect(GetCredentialStream());

            if (!isConnectionSuccessful)
            {
                if (connector.Status == ConnectStatus.NotConnected)
                {
                    Console.WriteLine($"Подключение отсутствует.\nПричина: {connector.Exception}");
                }
                else if (connector.Status == ConnectStatus.AuthorizationTimedOut)
                {
                    Console.WriteLine("Время на подключение закончилось");
                }

                Console.ReadLine();
 
                return;
            }

            Console.WriteLine("Подключились к Cloud App\n");

            const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

            if (connector.TryToGetSheetWithHeadAndKey(HttpUtils.GetSpreadsheetIdFromUri(uri),
                                                      HttpUtils.GetGidFromUri(uri),
                                                      "Head 1",
                                                      out SheetModel sheet))
            {
                var jsonEmptySheet = JsonSerialization.SerializeSheet(sheet, Formatting.Indented);

                sheet.AddRow(new List<string>() { "1", "2", "3", "4", "5" });
                sheet.AddRow(new List<string>() { "1", "2", "3", "4", "5" });
                sheet.AddRow(new List<string>() { "1", "2", "3", "4", "5" });
                sheet.AddRow(new List<string>() { "1", "2", "3", "4", "5" });
                sheet.AddRow(new List<string>() { "1", "2", "3", "4", "5" });
                connector.UpdateSheet(sheet);

                var jsonFullSheet = JsonSerialization.SerializeSheet(sheet, Formatting.Indented);

                var emptySheet = JsonSerialization.DeserializeSheet(jsonEmptySheet);
                sheet.Merge(emptySheet);
                connector.UpdateSheet(sheet);

                var fullSheet = JsonSerialization.DeserializeSheet(jsonFullSheet);
                sheet.Merge(fullSheet);
                connector.UpdateSheet(sheet);
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        private static Stream GetCredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
