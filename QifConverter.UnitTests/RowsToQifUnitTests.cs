using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace QifConverter.UnitTests
{
    [TestClass]
    public class RowsToQifUnitTests
    {
        [TestMethod]
        public void Only_Transactions_No_Initial_Amount_Ok()
        {
            // Given
            var rows = new List<Row>()
            {
                new Row{Date=DateTime.Now, Amount="-3,7", Label="Test"},
                new Row{Date=DateTime.Now, Amount="3,7", Label="Test"},
                new Row{Date=DateTime.Now, Amount="-3,7", Label="Test"}
            };

            // When
            var expected = "!Type:Bank\nD11/02/2021\nT-3,7\nMTest\n^\nD11/02/2021\nT3,7\nMTest\n^\nD11/02/2021\nT-3,7\nMTest\n^\n";
            var result = Program.RowsToQif(rows, 0, true);

            // Then
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Only_Transactions_Initial_Amount_Ok()
        {
            // Given
            var rows = new List<Row>()
            {
                new Row{Date=DateTime.Now, Amount="-3,7", Label="Test"},
                new Row{Date=DateTime.Now, Amount="3,7", Label="Test"},
                new Row{Date=DateTime.Now, Amount="-3,7", Label="Test"}
            };

            // When
            var expected = "!Type:Bank\nD11/02/2021\nT1000\nMInitial Amount\n^\nD11/02/2021\nT-3,7\nMTest\n^\nD11/02/2021\nT3,7\nMTest\n^\nD11/02/2021\nT-3,7\nMTest\n^\n";
            var result = Program.RowsToQif(rows, 1000, true);

            // Then
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void No_Transaction_No_Initial_Amount_Ok()
        {
            // Given
            var rows = new List<Row>()
            {
                new Row{Date=DateTime.Now, Amount="1000", Label="Test"},
                new Row{Date=DateTime.Now, Amount="1020", Label="Test"},
                new Row{Date=DateTime.Now, Amount="990", Label="Test"}
            };

            // When
            var expected = "!Type:Bank\nD11/02/2021\nT1000\nMTest\n^\nD11/02/2021\nT20\nMTest\n^\nD11/02/2021\nT-30\nMTest\n^\n";
            var result = Program.RowsToQif(rows, 0, false);

            // Then
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void No_Transaction_Initial_Amount_Ok()
        {
            // Given
            var rows = new List<Row>()
            {
                new Row{Date=DateTime.Now, Amount="1000", Label="Test"},
                new Row{Date=DateTime.Now, Amount="1020", Label="Test"},
                new Row{Date=DateTime.Now, Amount="990", Label="Test"}
            };

            // When
            var expected = "!Type:Bank\nD11/02/2021\nT1000\nMInitial Amount\n^\nD11/02/2021\nT0\nMTest\n^\nD11/02/2021\nT20\nMTest\n^\nD11/02/2021\nT-30\nMTest\n^\n";
            var result = Program.RowsToQif(rows, 1000, false);

            // Then
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Null_Label_Ok()
        {
            var rows = new List<Row>()
            {
                new Row{Date=DateTime.Now, Amount="1000", Label=""},
                new Row{Date=DateTime.Now, Amount="1020", Label=null},
                new Row{Date=DateTime.Now, Amount="990", Label=" "}
            };

            // When
            var expected = "!Type:Bank\nD11/02/2021\nT1000\nMInitial Amount\n^\nD11/02/2021\nT0\nMTransaction\n^\nD11/02/2021\nT20\nMTransaction\n^\nD11/02/2021\nT-30\nMTransaction\n^\n";
            var result = Program.RowsToQif(rows, 1000, false);

            // Then
            Assert.AreEqual(expected, result);
        }

    }
}
