using System;
using System.IO;
using GetGoogleSheetDataAPI;

namespace TestConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var connector = new Connector();
            bool isConnectionSuccessful = connector.TryToCreateConnect(CredentialStream());

            if (!isConnectionSuccessful)
            {
                if (connector.Status == ConnectStatus.NotConnected)
                {
                    Console.WriteLine("Подключение отсутствует");
                }
                else if (connector.Status == ConnectStatus.AuthorizationTimedOut)
                {
                    Console.WriteLine("Время на подключение закончилось");
                }
                
                return;
            }

            Console.WriteLine("Подключились к Cloud App");

            string url = "https://docs.google.com/spreadsheets/d/1dxPz9MEeJxfZkbAvZLE33YLIN5GaNS0Bvqzlp6rAiNk/edit#gid=0";

            if (connector.TryToGetSheet(url, out Sheet sheet))
            {
                PrintSheet(sheet);
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        private static void PrintSheet(Sheet sheet)
        {
            Console.WriteLine($"В листе {sheet.Rows.Count} строк");
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    Console.Write($"{cell.Value,3}|");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }

        private static Stream CredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
