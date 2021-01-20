using GetGoogleSheetDataAPI;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var connector = new Connector();
            // Попытка подключиться к приложению на Google Cloud Platform.
            // Метод вернёт флаг подключения и у экземпляра коннектора будет возможность
            // вызывать методы для получения листов google таблиц.
            bool isConnectionSuccessful = connector.TryToCreateConnect(GetCredentialStream());

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

            // замените режим для формирования таблицы с другой структурой
            // и другим доступом к этой структуре.
            var sheetMode = SheetMode.Head;

            string url = "https://docs.google.com/spreadsheets/d/1dxPz9MEeJxfZkbAvZLE33YLIN5GaNS0Bvqzlp6rAiNk/edit#gid=0";

            // Попытка получения листа по url.
            // Метод вернёт флаг получения листа и экземпляр типа Sheet.
            // Если данные из листа не были получены,
            // то экземпляр sheet будет пустым со статусом в виде строки,
            // говорящем о причине неудачи.
            // В противном случае экземпляр листа будет заполнен данными
            // и эти данные можно будет менять и через конннектор обновлять в листе google таблицы.
            if (connector.TryToGetSheet(url, sheetMode, out Sheet sheet))
            {
                PrintSheet(sheet);

                // Изменнения данных в экземпляре типа Sheet.
                // Не важен порядок изменения экземпляра листа,
                // его обновленние будет происходить в определённом порядке.
                AddRows(sheet);

                switch (sheetMode)
                {
                    case SheetMode.Simple:
                        ChangeRows(sheet);
                        break;
                    case SheetMode.Head:
                        ChangeRowsWithCellTitle(sheet);
                        break;
                    case SheetMode.HeadAndKey:
                        break;
                }

                DeleteRows(sheet);

                PrintSheet(sheet);

                // Метод для обновления данных в листе google таблицы на основе измененний
                // экземпляра типа Sheet.
                connector.UpdateSheet(sheet);
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        public static void AddRows(Sheet sheet)
        {
            // Пример добавления путой строки.
            sheet.AddRow();
            // Пример добавления строки часть которой будет заполнена пустыми строками.
            sheet.AddRow(new List<string>() { "123", "asd" });
            // Пример добавления строки где часть данных не попадёт в строку.
            sheet.AddRow(
                new List<string>()
                {
                    "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"
                }
            );
        }

        private static void ChangeRows(Sheet sheet)
        {
            // Можно было это делать в цикле например, но так проще.
            sheet.Rows[3].Cells[5].Value = "360";
            sheet.Rows[4].Cells[5].Value = "460";
            sheet.Rows[7].Cells[2].Value = "730";
            sheet.Rows[6].Cells[2].Value = "630";
            sheet.Rows[9].Cells[1].Value = "920";
        }

        private static void ChangeRowsWithCellTitle(Sheet sheet)
        {
            // Можно было это делать в цикле например, но так проще.
            sheet.Rows[3].Cells.Find(cell => cell.Title == "F").Value = "360";
            sheet.Rows[4].Cells.Find(cell => cell.Title == "F").Value = "460";
            sheet.Rows[7].Cells.Find(cell => cell.Title == "C").Value = "730";
            sheet.Rows[6].Cells.Find(cell => cell.Title == "C").Value = "630";
            sheet.Rows[9].Cells.Find(cell => cell.Title == "B").Value = "920";
        }

        public static void DeleteRows(Sheet sheet)
        {
            // Можно было это делать в цикле например, но так проще.
            sheet.DeleteRow(sheet.Rows[3]);
            sheet.DeleteRow(sheet.Rows[2]);
            sheet.DeleteRow(sheet.Rows[8]);
        }

        /// <summary>
        /// Просто метод для отображения данных таблицы в консоли
        /// </summary>
        /// <param name="sheet"></param>
        public static void PrintSheet(Sheet sheet)
        {
            Console.WriteLine($"Имя листа: {sheet.Title}. Имя таблицы: {sheet.SpreadsheetTitle}");
            Console.WriteLine($"В листе {sheet.Rows.Count} строк");
            
            foreach (var row in sheet.Rows)
            {
                Console.Write($"|{row.Number, 3}|");

                foreach (var cell in row.Cells)
                {
                    Console.Write($"|{cell.Value, 3}");
                }

                Console.Write($"| {row.Status}");
                Console.WriteLine();
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// Получение полномочий для подключения к приложению на Google Cloud Platform.
        /// </summary>
        /// <returns></returns>
        public static Stream GetCredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
