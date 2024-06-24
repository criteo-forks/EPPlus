using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class ExcelAddressInfoTests
    {
        [Test]
        public void ParseShouldThrowIfAddressIsNull()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                ExcelAddressInfo.Parse(null);
            });
        }

        [Test]
        public void ParseShouldSetWorksheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.That("Worksheet", Is.EqualTo(info.Worksheet));
        }

        [Test]
        public void WorksheetIsSpecifiedShouldBeTrueWhenWorksheetIsSupplied()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.That(info.WorksheetIsSpecified);
        }

        [Test]
        public void ShouldIndicateMultipleCellsWhenAddressContainsAColon()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.That(info.IsMultipleCells);
        }

        [Test]
        public void ShouldSetStartCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.That("A1", Is.EqualTo(info.StartCell));
        }

        [Test]
        public void ShouldSetEndCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.That("A2", Is.EqualTo(info.EndCell));
        }

        [Test]
        public void ParseShouldSetAddressOnSheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1:A2");
            Assert.That("A1:A2", Is.EqualTo(info.AddressOnSheet));
        }

        [Test]
        public void AddressOnSheetShouldBeSameAsAddressIfNoWorksheetIsSpecified()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.That("A1:A2", Is.EqualTo(info.AddressOnSheet));
        }
    }
}
