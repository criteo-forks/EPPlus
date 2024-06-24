using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
    [TestFixture]
    public class SqRefUtilityTests
    {
        [Test]
        public void SqRefUtility_ToSqRefAddress_ShouldThrowIfAddressIsNullOrEmpty()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                SqRefUtility.ToSqRefAddress(null);
            });
        }

        [Test]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommas()
        {
            // Arrange
            var address = "A1, A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.That("A1 A2:A3", Is.EqualTo(result));
        }


        [Test]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommasAndInsertSpaceIfNecesary()
        {
            // Arrange
            var address = "A1,A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.That("A1 A2:A3", Is.EqualTo(result));
        }

        [Test]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveMultipleSpaces()
        {
            // Arrange
            var address = "A1,        A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.That("A1 A2:A3", Is.EqualTo(result));
        }

        [Test]
        public void SqRefUtility_FromSqRefAddress_ShouldReplaceSpaceWithComma()
        {
            // Arrange
            var address = "A1 A2";

            // Act
            var result = SqRefUtility.FromSqRefAddress(address);

            // Assert
            Assert.Equals("A1,A2", result);
        }
    }
}
