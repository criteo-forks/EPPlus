using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class ExcelTimeTests
    {
        private ExcelTime _time;
        private readonly decimal SecondsPerHour = 3600;
       // private readonly decimal HoursPerDay = 24;
        private readonly decimal SecondsPerDay = 3600 * 24;

        private decimal Round(decimal value)
        {
            return Math.Round(value, ExcelTime.NumberOfDecimals);
        }

        [SetUp]
        public void Setup()
        {
            _time = new ExcelTime();
        }

        [TearDown]
        public void Cleanup()
        {
            _time = null;
        }

        [Test]
        public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsLessThan0()
        {
            Assert.Throws<ArgumentException>(() => { 
                new ExcelTime(-1);
            });
        }

        [Test]
        public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsEqualToOrGreaterThan1()
        {
            Assert.Throws<ArgumentException>(() => { 
                new ExcelTime(1);
            });
        }

        [Test]
        public void ExcelTimeTests_Hour_ShouldThrowIfNegativeValue()
        {
            Assert.Throws<InvalidOperationException>(() => { 
                _time.Hour = -1;
            });
        }

        [Test]
        public void ExcelTimeTests_Minute_ShouldThrowIfNegativeValue()
        {
            Assert.Throws<InvalidOperationException>(() => { 
                _time.Minute = -1;
            });
        }

        [Test]
        public void ExcelTimeTests_Minute_ShouldThrowIValueIsGreaterThan59()
        {
            Assert.Throws<InvalidOperationException>(() => { 
                _time.Minute = 60;
            });
        }

        [Test]
        public void ExcelTimeTests_Second_ShouldThrowIfNegativeValue()
        {
            Assert.Throws<InvalidOperationException>(() => {
                _time.Second = -1;
            });
        }

        [Test]
        public void ExcelTimeTests_Second_ShouldThrowIValueIsGreaterThan59()
        {
            Assert.Throws<InvalidOperationException>(() => {
                _time.Second = 60;
            });
        }

        [Test]
        public void ExcelTimeTests_ToExcelTime_HourIsSet()
        {
            // Act
            _time.Hour = 1;
            
            // Assert
            Assert.That(Round(SecondsPerHour/SecondsPerDay), Is.EqualTo(_time.ToExcelTime()));
        }

        [Test]
        public void ExcelTimeTests_ToExcelTime_MinuteIsSet()
        {
            // Arrange
            decimal expected = SecondsPerHour + (20M * 60M);
            // Act
            _time.Hour = 1;
            _time.Minute = 20;

            // Assert
            Assert.That(Round(expected/SecondsPerDay), Is.EqualTo(_time.ToExcelTime()));
        }

        [Test]
        public void ExcelTimeTests_ToExcelTime_SecondIsSet()
        {
            // Arrange
            decimal expected = SecondsPerHour + (20M * 60M) + 10M;
            // Act
            _time.Hour = 1;
            _time.Minute = 20;
            _time.Second = 10;

            // Assert
            Assert.That(Round(expected / SecondsPerDay), Is.EqualTo(_time.ToExcelTime()));
        }

        [Test]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetHour()
        {
            // Arrange
            decimal value = 3660M/(decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.That(1, Is.EqualTo(time.Hour));
        }

        [Test]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetMinute()
        {
            // Arrange
            decimal value = 3660M / (decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.That(1, Is.EqualTo(time.Minute));
        }

        [Test]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetSecond()
        {
            // Arrange
            decimal value = 3662M / (decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.That(2, Is.EqualTo(time.Second));
        }
    }
}
