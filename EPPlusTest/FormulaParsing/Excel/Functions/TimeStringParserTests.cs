using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class TimeStringParserTests
    {
        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        [Test]
        public void CanParseShouldHandleValid24HourPatterns()
        {
            var parser = new TimeStringParser();
            Assert.That(parser.CanParse("10:12:55"), "Could not parse 10:12:55");
            Assert.That(parser.CanParse("22:12:55"), "Could not parse 13:12:55");
            Assert.That(parser.CanParse("13"), "Could not parse 13");
            Assert.That(parser.CanParse("13:12"), "Could not parse 13:12");
        }

        [Test]
        public void CanParseShouldHandleValid12HourPatterns()
        {
            var parser = new TimeStringParser();
            Assert.That(parser.CanParse("10:12:55 AM"), "Could not parse 10:12:55 AM");
            Assert.That(parser.CanParse("9:12:55 PM"), "Could not parse 9:12:55 PM");
            Assert.That(parser.CanParse("7 AM"), "Could not parse 7 AM");
            Assert.That(parser.CanParse("4:12 PM"), "Could not parse 4:12 PM");
        }

        [Test]
        public void ParseShouldIdentifyPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:12:55");
            Assert.That(GetSerialNumber(10, 12, 55), Is.EqualTo(result));
        }

        [Test]
        public void ParseShouldThrowExceptionIfSecondIsOutOfRange()
        {
            Assert.Throws<FormatException>(() =>
            {
                var parser = new TimeStringParser();
                var result = parser.Parse("10:12:60");
            });
            
        }

        [Test]
        public void ParseShouldThrowExceptionIfMinuteIsOutOfRange()
        {
            Assert.Throws<FormatException>(() =>
            {
                var parser = new TimeStringParser();
                var result = parser.Parse("10:60:55");
            });
        }

        [Test]
        public void ParseShouldIdentify12HourAMPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:12:55 AM");
            Assert.That(GetSerialNumber(10, 12, 55), Is.EqualTo(result));
        }

        [Test]
        public void ParseShouldIdentify12HourPMPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:12:55 PM");
            Assert.That(GetSerialNumber(22, 12, 55), Is.EqualTo(result));
        }
    }
}
