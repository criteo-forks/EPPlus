﻿using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;
namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class SubtotalTests : FormulaParserTestBase
    {
        private ExcelWorksheet _worksheet;
        private ExcelPackage _package;

        [SetUp]
        public void Setup()
        {
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("Test");
            _parser = _worksheet.Workbook.FormulaParser;
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Count()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(2, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_CountA()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(3, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Max()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(4, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Min()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(5, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Product()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(6, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Stdev()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(7, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(7, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            result = Math.Round((double)result, 9);
            Assert.That(0.707106781d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_StdevP()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(8, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(0.5d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Sum()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_Var()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(9d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNotIncludeSubtotalChildren_VarP()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(10, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.That(0.5d, Is.EqualTo(result));
        }
    }
}
