using SynSys.GSpreadsheetEasyAccess.Data;
using System;
using System.Collections.Generic;

namespace Printing
{
    public static class Printer
    {
        private const int _firstColumnWith = 3;
        private static int _columnWith;

        public static void PrintSheet(AbstractSheet sheet, string status)
        {
            // TODO написать универсальный принтер без хардкода ширины столбцов.
            _columnWith = 9;

            PrintDescription(sheet, status);
            PrintSeparateRow(sheet.Head);
            PrintHead(sheet.Head);
            PrintSeparateRow(sheet.Head);
            PrintBody(sheet.Rows);
            PrintSeparateRow(sheet.Head);
        }

        private static void PrintDescription(AbstractSheet sheet, string status)
        {
            ColorizeText(ChangeStatus.Original);
            Console.WriteLine(
                "\n" +
               $"Status:            {status}\n" +
               $"Spreadsheet Title: {sheet.SpreadsheetTitle}\n" +
               $"Sheet Title:       {sheet.Title}\n" +
               $"Number of columns: {sheet.Head.Count}\n" +
               $"Number of rows:    {sheet.Rows.Count}\n"
            );
        }

        private static void PrintHead(List<AbstractColumn> head)
        {
            PrintColumnsStatuses(head);
            PrintColumnsNumbers(head);
            PrintSeparateRow(head);
            PrintColumnTitles(head);
        }

        private static void PrintColumnsStatuses(List<AbstractColumn> head)
        {
            Console.Write("".PadRight(_firstColumnWith));

            foreach (AbstractColumn column in head)
            {
                ColorizeText(column.Status);
                Console.Write($"|{column.Status}".PadRight(_columnWith));
            }

            PrintRowEnding(ChangeStatus.Original, string.Empty);
        }

        private static void PrintColumnsNumbers(List<AbstractColumn> head)
        {
            Console.Write("".PadRight(_firstColumnWith));

            foreach (AbstractColumn column in head)
            {
                ColorizeText(column.Status);
                Console.Write($"|{column.Number}".PadRight(_columnWith));
            }

            PrintRowEnding(ChangeStatus.Original, string.Empty);
        }

        private static void PrintSeparateRow(List<AbstractColumn> head)
        {
            const string lineCross = "+";
            var delimiter = new string('-', _columnWith - 1);

            ColorizeText(ChangeStatus.Original);
            Console.Write("".PadRight(_firstColumnWith));

            foreach (AbstractColumn column in head)
            {
                ColorizeText(column.Status);
                Console.Write($"{lineCross}{delimiter}");
            }

            PrintRowEnding(ChangeStatus.Original, string.Empty, lineCross);
        }

        private static void PrintColumnTitles(List<AbstractColumn> head)
        {
            Console.Write("".PadRight(_firstColumnWith));

            foreach (AbstractColumn column in head)
            {
                string value = column.Title;

                if (string.IsNullOrWhiteSpace(column.Title))
                {
                    value = "-";
                }

                ColorizeText(column.Status);
                Console.Write($"|{value}".PadRight(_columnWith));
            }

            PrintRowEnding(ChangeStatus.Original, string.Empty);
        }

        private static void PrintBody(List<Row> rows)
        {
            foreach (Row row in rows)
            {
                ColorizeText(row.Status);
                Console.Write($"{row.Number}".PadLeft(_firstColumnWith));

                foreach (Cell cell in row.Cells)
                {
                    Console.Write($"|{cell.Value}".PadRight(_columnWith));
                }

                PrintRowEnding(row.Status, row.Status.ToString());
            }
        }

        private static void PrintRowEnding(ChangeStatus changeStatus, string statusMessage, string symbol = "|")
        {
            ColorizeText(changeStatus);
            Console.Write($"{symbol} {statusMessage}");
            Console.WriteLine();
        }

        private static void ColorizeText(ChangeStatus status)
        {
            if (status == ChangeStatus.ToAppend)
            {
                Console.ForegroundColor = ConsoleColor.Green;
            }
            else if (status == ChangeStatus.ToInsert)
            {
                Console.ForegroundColor = ConsoleColor.DarkGreen;
            }
            else if (status == ChangeStatus.ToChange)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
            }
            else if (status == ChangeStatus.ToDelete)
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
    }
}
