using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SynSys.GSpreadsheetEasyAccess.Application;

namespace SynSys.GSpreadsheetEasyAccess.Tests
{
    [TestClass]
    public class GoogleSheetUriTests
    {
        [TestMethod]
        public void RightGoogleSheetUriTest()
        {
            // arrange
            var spreadsheetId = "11Q_dOgIkMIX96q7e13RrmeionRXOW1mgzG98dxXOCFU";
            var gid = "1108304943";
            var uri = $"https://docs.google.com/d/{spreadsheetId}/edit#gid={gid}";

            // act
            var guri = new GoogleSheetUri(uri);

            // assert
            Assert.AreEqual(spreadsheetId, guri.SpreadsheetId);
            Assert.AreEqual(gid, guri.SheetId);
        }

        [TestMethod]
        public void RandomTextTest()
        {
            // arrange
            var uri = "asdf";

            // assert
            Assert.ThrowsException<UriFormatException>(() => new GoogleSheetUri(uri));
        }

        [TestMethod]
        public void UriWithoutGoogleDocsDomainTest()
        {
            // arrange
            var uri = "https://spreadsheets/d/11Q_dOgIkMIX96q7e13RrmeionRXOW1mgzG98dxXOCFU/edit#gid=1108304943";

            // assert
            Assert.ThrowsException<UriFormatException>(() => new GoogleSheetUri(uri));
        }

        [TestMethod]
        public void UriWithoutSpreadsheetAffiliationTest()
        {
            // arrange
            var uri = "https://docs.google.com/d/11Q_dOgIkMIX96q7e13RrmeionRXOW1mgzG98dxXOCFU/edit#gid=1108304943";

            // assert
            Assert.ThrowsException<UriFormatException>(() => new GoogleSheetUri(uri));
        }

        [TestMethod]
        public void UriWithoutSpreadsheetIdTest()
        {
            // arrange
            var uri = "https://docs.google.com/spreadsheets/edit#gid=1108304943";

            // assert
            Assert.ThrowsException<UriFormatException>(() => new GoogleSheetUri(uri));
        }

        [TestMethod]
        public void UriWithoutSheetIdTest()
        {
            // arrange
            var uri = "https://docs.google.com/spreadsheets/d/11Q_dOgIkMIX96q7e13RrmeionRXOW1mgzG98dxXOCFU/edit#";

            // assert
            Assert.ThrowsException<UriFormatException>(() => new GoogleSheetUri(uri));
        }
    }
}