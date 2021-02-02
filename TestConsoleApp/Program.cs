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

            // замените url адрес на свой для тестов.
            string url = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0";

            // Попытка получения листа по url.
            // Метод вернёт флаг получения листа и экземпляр типа Sheet.
            // Если данные из листа не были получены,
            // то экземпляр sheet будет пустым со статусом в виде строки,
            // говорящем о причине неудачи.
            // В противном случае экземпляр листа будет заполнен данными
            // и эти данные можно будет менять и через конннектор обновлять в листе google таблицы.

            //if (connector.TryToGetSimpleSheet(url, out Sheet sheet))
            //if (connector.TryToGetSheetWithHead(url, out Sheet sheet))
            if (connector.TryToGetSheetWithHeadAndKey(url, "A", out Sheet sheet))
            {
                PrintSheet(sheet, "Первое получение данных");
                ChangeSheet(connector, sheet);
                // Данный метод вызывается второй раз чтобы показать,
                // что с текущим экземпляром листа можно работать и после обновления.
                // Его структура будет соответствовать google таблице.
                ChangeSheet(connector, sheet);
            }
            else
            {
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        private static void ChangeSheet(Connector connector, Sheet sheet)
        {
            // Изменнения данных в экземпляре типа Sheet.
            // Не важен порядок изменения экземпляра листа,
            // его обновленние будет происходить в определённом порядке.
            AddRows(sheet);

            switch (sheet.Mode)
            {
                case SheetMode.Simple:
                    ChangeRows(sheet);
                    break;
                case SheetMode.Head:
                    ChangeRowsWithCellTitle(sheet);
                    break;
                case SheetMode.HeadAndKey:
                    ChangeRowsWithTitle(sheet);
                    break;
            }

            DeleteRows(sheet);

            PrintSheet(sheet, "Данные до обновления в google");

            // Метод для обновления данных в листе google таблицы на основе измененний
            // экземпляра типа Sheet.
            connector.UpdateSheet(sheet);

            PrintSheet(sheet, "Данные после обновления в google");
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
            // Можно было это делать в цикле, но так проще.
            sheet.Rows[3].Cells[5].Value = "360";
            sheet.Rows[4].Cells[5].Value = "460";
            sheet.Rows[7].Cells[2].Value = "730";
            sheet.Rows[6].Cells[2].Value = "630";
            sheet.Rows[9].Cells[1].Value = "920";
        }

        private static void ChangeRowsWithCellTitle(Sheet sheet)
        {
            // Можно было это делать в цикле, но так проще.
            sheet.Rows[2].Cells.Find(cell => cell.Title == "F").Value = "360";
            sheet.Rows[3].Cells.Find(cell => cell.Title == "F").Value = "460";
            sheet.Rows[6].Cells.Find(cell => cell.Title == "C").Value = "730";
            sheet.Rows[5].Cells.Find(cell => cell.Title == "C").Value = "630";
            sheet.Rows[8].Cells.Find(cell => cell.Title == "B").Value = "920";
        }

        private static void ChangeRowsWithTitle(Sheet sheet)
        {
            sheet.Rows
                 .Find(row => row.Key.Value == "61")
                 .Cells
                 .Find(cell => cell.Title == "F")
                 .Value = "660";
        }

        public static void DeleteRows(Sheet sheet)
        {
            // Можно было это делать в цикле, но так проще.
            sheet.DeleteRow(sheet.Rows[3]);
            sheet.DeleteRow(sheet.Rows[2]);
            sheet.DeleteRow(sheet.Rows[8]);
        }

        /// <summary>
        /// Просто метод для отображения данных таблицы в консоли
        /// </summary>
        /// <param name="sheet"></param>
        public static void PrintSheet(Sheet sheet, string status)
        {
            Console.WriteLine(
                "\n" +
                $"Статус листа:     {status}\n" +
                $"Имя таблицы:      {sheet.SpreadsheetTitle}\n" +
                $"Имя листа:        {sheet.Title}\n" +
                $"Количество строк: {sheet.Rows.Count}"
            );

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
