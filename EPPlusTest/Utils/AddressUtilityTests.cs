using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.Utils;
using OfficeOpenXml;

namespace EPPlusTest.Utils
{
    [TestFixture]
    public class AddressUtilityTests
    {
        [Test]
        public void ParseForEntireColumnSelections_ShouldAddMaxRows()
        {
            // Arrange
            var address = "A:A";

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.That("A1:A" + ExcelPackage.MaxRows, Is.EqualTo(result));
        }

        [Test]
        public void ParseForEntireColumnSelections_ShouldAddMaxRowsOnColumnsWithMultipleLetters()
        {
            // Arrange
            var address = "AB:AC";

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.That("AB1:AC" + ExcelPackage.MaxRows, Is.EqualTo(result));
        }

        [Test]
        public void ParseForEntireColumnSelections_ShouldHandleMultipleRanges()
        {
            // Arrange
            var address = "A:A B:B";
            var expected = string.Format("A1:A{0} B1:B{0}", ExcelPackage.MaxRows);

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.That(expected, Is.EqualTo(result));
        }
    }
}
