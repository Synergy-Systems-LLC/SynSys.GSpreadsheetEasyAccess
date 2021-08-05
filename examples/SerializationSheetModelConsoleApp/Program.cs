using Newtonsoft.Json;
using SynSys.GSpreadsheetEasyAccess;
using System;
using System.Collections.Generic;
using System.IO;

namespace SerializationSheetModelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Запуск {nameof(SerializationSheetModelConsoleApp)}\n");

            var connector = new Connector()
            {
                ApplicationName = "Get google sheet data",
            };
 
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

            if (connector.TryToGetSheetWithHeadAndKey(uri, "Head 1", out SheetModel sheet))
            {
                PrintSheet(sheet, "Первое получение данных");

                string jsonSheet = JsonSerialization.SerializeSheet(sheet, Formatting.Indented);
                Console.WriteLine($"\nЛист после сериализации\n{jsonSheet}");

                SheetModel deserializedSheet = JsonSerialization.DeserializeSheet(jsonSheet);
                PrintSheet(deserializedSheet, "Лист после десериализации");

                ChangeSheet(sheet);
                PrintSheet(sheet, "Изменённые данные до обновления в google");

                connector.UpdateSheet(sheet);
                PrintSheet(sheet, "Данные после обновления в google");
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        private static void ChangeSheet(SheetModel sheet)
        {
            AddRows(sheet);
            ChangeRows(sheet);
            DeleteRows(sheet);
        }

        private static void AddRows(SheetModel sheet)
        {
            sheet.AddRow();
            sheet.AddRow(new List<string> { "123", "456", "789" });
        }

        private static void ChangeRows(SheetModel sheet)
        {
            sheet.Rows.Find(row => row.Key.Value == "41")
                 .Cells.Find(cell => cell.Title == "Head 6")
                 .Value = "460";
            sheet.Rows.Find(row => row.Key.Value == "71")
                 .Cells.Find(cell => cell.Title == "Head 2")
                 .Value = "720";
        }

        private static void DeleteRows(SheetModel sheet)
        {
            sheet.DeleteRow(sheet.Rows[2]);
            sheet.DeleteRow(sheet.Rows[8]);
        }

        private static void PrintSheet(SheetModel sheet, string status)
        {
            Console.WriteLine(
                "\n" +
               $"Статус листа:     {status}\n" +
               $"Имя таблицы:      {sheet.SpreadsheetTitle}\n" +
               $"Имя листа:        {sheet.Title}\n" +
               $"Количество строк: {sheet.Rows.Count}"
            );

            foreach (Row row in sheet.Rows)
            {
                Console.Write($"|{row.Number, 3}|");

                foreach (Cell cell in row.Cells)
                {
                    Console.Write($"|{cell.Value, 7}");
                }

                Console.Write($"| {row.Status}");
                Console.WriteLine();
            }
        }

        private static Stream GetCredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
