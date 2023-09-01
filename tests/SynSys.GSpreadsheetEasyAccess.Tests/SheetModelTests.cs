using Microsoft.VisualStudio.TestTools.UnitTesting;
using SynSys.GSpreadsheetEasyAccess.Data;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess.Tests
{
    [TestClass]
    public class SheetModelTests
    {
        SheetModel sheet;

        [TestInitialize]
        public void Init()
        {
            var data = new List<IList<object>>()
            {
                new List<object>() { "Head 1", "Head 2", "Head 3" },
                new List<object>() { "qwer", "tyui", "op[]" },
                new List<object>() { "asdf", "ghjk", "l;'" },
            };

            sheet = new SheetModel
            {
                Mode = SheetMode.HeadAndKey,
                KeyName = "Head 1",
                Gid = 0,
                Title = "TestTitle",
                SpreadsheetId = "0000000000",
                SpreadsheetTitle = "TestSpreadsheetTitle"
            };

            sheet.Fill(data);
        }

        /// <summary>
        /// Тест проверяет корректность статусов всех строк.
        /// Clean должен изменить статус каждой строки на RowStatus.ToDelete.
        /// </summary>
        [TestMethod]
        public void Clean_RowsStatus()
        {
            // arrange
            var expectedRowStatus = ChangeStatus.ToDelete;

            // Действия для наличия различных статусов с листе.
            sheet.Rows[0].Cells[0].Value = "1234";
            sheet.AddRow();

            // act
            sheet.Clean();

            // assert
            Assert.IsTrue(sheet.Rows.TrueForAll(r => r.Status == expectedRowStatus));
        }

        /// <summary>
        /// Тест проверяет корректность количества строк.
        /// Clean должен физически удалять строки со статусом RowStatus.ToAppend из sheet.Rows.
        /// </summary>
        [TestMethod]
        public void Clean_RowsAmount()
        {
            // arrange
            sheet.AddRow();
            sheet.AddRow( new List<string>() { "12", "34", "56" });

            var expectedRowsAmount = sheet.Rows.FindAll(r => r.Status != ChangeStatus.ToAppend).Count;

            // act
            sheet.Clean();

            // assert
            Assert.AreEqual(
                expectedRowsAmount,
                sheet.Rows.Count,
                $"\nexpect: {expectedRowsAmount}" +
                $"\nactual: {sheet.Rows.Count}"
            );
        }

        /// <summary>
        /// Тест проверяет корректность определения количества потеряных заголовков.
        /// </summary>
        [TestMethod]
        public void DoesNotContainsSomeHeaders_AllHeadersExists()
        {
            // arrage
            string[] requiredHeaders = { "Head 1", "Head 2", "Head 3", "Head 4" };

            // assert
            Assert.ThrowsException<InvalidSheetHeadException>(() => sheet.CheckHead(requiredHeaders));
        }

        /// <summary>
        /// Тест проверяет корректное удаление добавляемой строки.<br/>
        /// Строка со статусом RowStatus.ToAppend должна физически удалиться из списка строк.
        /// Все последующие должны уменьшить свой номер на 1.
        /// </summary>
        [TestMethod]
        public void DeleteRow_DeleteFirstAppendRow()
        {
            // arrage
            sheet.AddRow();
            sheet.AddRow();
            sheet.AddRow( new List<string>() { "12", "34", "56" });

            var startIndex = 2;
            var count = sheet.Rows.Count - 1;
            var expectedRangeNumbers = Enumerable.Range(startIndex, count).ToList();
            var appendRow = sheet.Rows.Find(r => r.Status == ChangeStatus.ToAppend);

            // act
            sheet.DeleteRow(appendRow);

            // assert
            var actualRangeNumbers = sheet.Rows.Select(r => r.Number).ToList();
            CollectionAssert.AreEqual(
                expectedRangeNumbers,
                actualRangeNumbers,
                $"\nexpect: {string.Join(" ", expectedRangeNumbers)}" +
                $"\nactual: {string.Join(" ", actualRangeNumbers)}\n"
            );
        }

        /// <summary>
        /// Тест проверяет корректное изменение статуса строки.<br/>
        /// Строки со статусами RowStatus.Original и RowStatus.ToChange
        /// должны измениться на RowStatus.ToDelete.
        /// </summary>
        [TestMethod]
        public void DeleteRow_FindRowsWithDeleteRowStatus()
        {
            // arrage
            sheet.Rows[0].Cells[0].Value = "1234";

            var originalRow = sheet.Rows.Find(r => r.Status == ChangeStatus.Original);
            var changeRow = sheet.Rows.Find(r => r.Status == ChangeStatus.ToChange);
            var expectedRowStatus = ChangeStatus.ToDelete;

            // act
            sheet.DeleteRow(originalRow);
            sheet.DeleteRow(changeRow);

            // assert
            Assert.AreEqual(
                expectedRowStatus,
                originalRow.Status,
                $"\nactual: {originalRow.Status}"
            );
            Assert.AreEqual(
                expectedRowStatus,
                changeRow.Status,
                $"\nactual: {changeRow.Status}"
            );
        }
    }
}
