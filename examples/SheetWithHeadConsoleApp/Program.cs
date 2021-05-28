using SynSys.GSpreadsheetEasyAccess;
using System;
using System.Collections.Generic;
using System.IO;

namespace SheetWithHeadConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Запуск SheetWithHeadConsoleApp\n");

            // Инициализируем коннектор и настраиваем его если нужно.
            var connector = new Connector()
            {
                // Установка данных свойств опциональна.
                ApplicationName = "Get google sheet data",
                CancellationSeconds = 20,
            };
 
            // Попытка подключиться к приложению на Google Cloud Platform.
            // Метод вернёт флаг подключения и у экземпляра коннектора будет возможность
            // вызывать методы для получения листов гугл таблиц.
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

            // Замените url адрес на свой для тестов.
            // Если вы воспользовались данным uri адресом, то после успешного завершения работы программы,
            // скопируйте содержимое из листа "Copy of TestSheet" в лист "TestSheet".
            const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

            // Попытка получения листа по uri.
            // Метод вернёт флаг получения листа и экземпляр типа SheetModel.
            // Если данные из листа не были получены,
            // то экземпляр SheetModel будет пустым со статусом в виде строки, говорящем о причине неудачи.
            // В противном случае экземпляр листа будет заполнен данными.
            // Эти данные можно менять и через конннектор обновлять в листе гугл таблицы.
            //if (connector.TryToGetSimpleSheet(uri, out SheetModel sheet))
            //if (connector.TryToGetSimpleSheet(HttpUtils.GetSpreadsheetIdFromUri(uri),
            //                                  "TestSheet", out SheetModel sheet))
            if (connector.TryToGetSheetWithHead(HttpUtils.GetSpreadsheetIdFromUri(uri),
                                                HttpUtils.GetGidFromUri(uri), out SheetModel sheet))
            {
                PrintSheet(sheet, "Первое получение данных");

                ChangeSheet(sheet);
                PrintSheet(sheet, "Изменённые данные до обновления в google");

                // Метод для обновления данных в листе google таблицы на основе измененний
                // экземпляра типа Sheet.
                connector.UpdateSheet(sheet);
                PrintSheet(sheet, "Данные после обновления в google");

                // Данный метод вызывается второй раз чтобы показать,
                // что с текущим экземпляром листа можно работать и после обновления.
                // Его структура будет соответствовать google таблице.
                ChangeSheet(sheet);
                PrintSheet(sheet, "Изменённые данные до обновления в google");

                connector.UpdateSheet(sheet);
                PrintSheet(sheet, "Данные после обновления в google");
            }
            else
            {
                // По какой-то причине не получилось получить таблицу.
                // Причина будет указана в sheet.Status.
                Console.WriteLine(sheet.Status);
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Изменения данных в экземпляре типа Sheet.
        /// Не важен порядок изменения экземпляра листа,
        /// его обновленние будет происходить в определённом порядке в коннекторе.
        /// </summary>
        /// <param name="sheet"></param>
        private static void ChangeSheet(SheetModel sheet)
        {
            AddRows(sheet);
            ChangeRows(sheet);
            DeleteRows(sheet);
        }

        private static void AddRows(SheetModel sheet)
        {
            // Добавления пустой строки.
            sheet.AddRow();

            // Добавления строки часть которой будет заполнена пустыми строками.
            sheet.AddRow(new List<string>() { "123", "asd" });

            // Добавления строки где часть данных не попадёт в строку, а именно "k" и "l".
            // Потому что в "TestSheet" 10 столбцов.
            sheet.AddRow(
                new List<string>()
                {
                    "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"
                }
            );
        }

        private static void ChangeRows(SheetModel sheet)
        {
            // Данный пример не учитывает отсутствие ячеек с выбранными Title
            sheet.Rows[2].Cells.Find(cell => cell.Title == "F").Value = "360";
            sheet.Rows[3].Cells.Find(cell => cell.Title == "F").Value = "460";
            sheet.Rows[6].Cells.Find(cell => cell.Title == "C").Value = "730";
            sheet.Rows[5].Cells.Find(cell => cell.Title == "C").Value = "630";
            sheet.Rows[8].Cells.Find(cell => cell.Title == "B").Value = "920";
        }

        private static void DeleteRows(SheetModel sheet)
        {
            // Данный пример не учитывает отсутствие нужных индексов
            sheet.DeleteRow(sheet.Rows[3]);
            sheet.DeleteRow(sheet.Rows[2]);
            sheet.DeleteRow(sheet.Rows[8]);
        }

        /// <summary>
        /// Просто метод для отображения данных таблицы в консоли
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="status"></param>
        private static void PrintSheet(SheetModel sheet, string status)
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
        private static Stream GetCredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
