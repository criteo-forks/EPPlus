﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class InformationFunctionsTests
    {
        private ParsingContext _context;

        [SetUp]
        public void Setup()
        {
            _context = ParsingContext.Create();
        }

        [Test]
        public void IsBlankShouldReturnTrueIfFirstArgIsNull()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(new object[]{null});
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsBlankShouldReturnTrueIfFirstArgIsEmptyString()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(string.Empty);
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsNumberShouldReturnTrueWhenArgIsNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs(1d);
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsNumberShouldReturnfalseWhenArgIsNonNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs("1");
            var result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void IsErrorShouldReturnTrueIfArgIsAnErrorCode()
        {
            var args = FunctionsHelper.CreateArgs(ExcelErrorValue.Parse("#DIV/0!"));
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsErrorShouldReturnFalseIfArgIsNotAnError()
        {
            var args = FunctionsHelper.CreateArgs("A", 1);
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void IsTextShouldReturnTrueWhenFirstArgIsAString()
        {
            var args = FunctionsHelper.CreateArgs("1");
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsTextShouldReturnFalseWhenFirstArgIsNotAString()
        {
            var args = FunctionsHelper.CreateArgs(1);
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void IsNonTextShouldReturnFalseWhenFirstArgIsAString()
        {
            var args = FunctionsHelper.CreateArgs("1");
            var func = new IsNonText();
            var result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void IsNonTextShouldReturnTrueWhenFirstArgIsNotAString()
        {
            var args = FunctionsHelper.CreateArgs(1);
            var func = new IsNonText();
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsOddShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(3.123);
            var func = new IsOdd();
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsEvenShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(4.123);
            var func = new IsEven();
            var result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IsLogicalShouldReturnCorrectResult()
        {
            var func = new IsLogical();

            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);

            args = FunctionsHelper.CreateArgs("true");
            result = func.Execute(args, _context);
            Assert.That(!(bool)result.Result);

            args = FunctionsHelper.CreateArgs(false);
            result = func.Execute(args, _context);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void NshouldReturnCorrectResult()
        {
            var func = new N();

            var args = FunctionsHelper.CreateArgs(1.2);
            var result = func.Execute(args, _context);
            Assert.That(1.2, Is.EqualTo(result.Result));

            args = FunctionsHelper.CreateArgs("abc");
            result = func.Execute(args, _context);
            Assert.That(0d, Is.EqualTo(result.Result));

            args = FunctionsHelper.CreateArgs(true);
            result = func.Execute(args, _context);
            Assert.That(1d, Is.EqualTo(result.Result));

            var errorCode = ExcelErrorValue.Create(eErrorType.Value);
            args = FunctionsHelper.CreateArgs(errorCode);
            result = func.Execute(args, _context);
            Assert.That(errorCode, Is.EqualTo(result.Result));
        }
    }
}
