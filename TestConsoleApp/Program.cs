using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

                AddRows(sheet);
                ChangeRows(sheet);
                DeleteRows(sheet);

                PrintSheet(sheet);

                connector.UpdateSheet(sheet);
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        private static void ChangeRows(Sheet sheet)
        {
            sheet.Rows[3].Cells[5].Value = "360";
            sheet.Rows[4].Cells[5].Value = "460";
            sheet.Rows[7].Cells[2].Value = "730";
            sheet.Rows[6].Cells[2].Value = "630";
            sheet.Rows[9].Cells[1].Value = "920";
        }

        private static void DeleteRows(Sheet sheet)
        {
            sheet.DeleteRow(sheet.Rows[3]);
            sheet.DeleteRow(sheet.Rows[2]);
            sheet.DeleteRow(sheet.Rows[8]);
        }

        private static void AddRows(Sheet sheet)
        {
            sheet.AddRow();
            sheet.AddRow(new List<string>() { "123", "asd" });
            sheet.AddRow(
                new List<string>()
                {
                    "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"
                }
            );
        }

        private static void PrintSheet(Sheet sheet)
        {
            Console.WriteLine($"В листе {sheet.Rows.Count} строк");
            
            foreach (var row in sheet.Rows)
            {
                Console.Write($"|{row.Number, 3}|");

                foreach (var cell in row.Cells)
                {
                    Console.Write($"|{cell.Value,3}");
                }

                Console.Write($"| {row.Status}");
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
