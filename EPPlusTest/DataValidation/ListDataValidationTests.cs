using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class ListDataValidationTests : ValidationTestBase
    {
        private IExcelDataValidationList _validation;

        [SetUp]
        public void Setup()
        {
            SetupTestData();
            _validation = _sheet.Workbook.Worksheets[1].DataValidations.AddListValidation("A1");
        }

        [TearDown]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [Test]
        public void ListDataValidation_FormulaIsSet()
        {
            Assert.That(_validation.Formula, Is.Not.Null);
        }

        [Test]
        public void ListDataValidation_WhenOneItemIsAddedCountIs1()
        {
            // Act
            _validation.Formula.Values.Add("test");

            // Assert
            Assert.That(1, Is.EqualTo(_validation.Formula.Values.Count));
        }

        [Test]
        public void ListDataValidation_ShouldThrowWhenNoFormulaOrValueIsSet()
        {
            Assert.Throws<InvalidOperationException>(() => { 
                _validation.Validate();
            });
        }
    }
}
