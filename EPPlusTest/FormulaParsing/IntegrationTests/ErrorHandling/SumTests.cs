using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ErrorHandling
{
    /// <summary>
    /// Summary description for SumTests
    /// </summary>
    [TestFixture]
    [Explicit]
    public class SumTests : FormulaErrorHandlingTestBase
    {
        [SetUp]
        public void ClassInitialize()
        {
            BaseInitialize();
        }

        [TearDown]
        public void ClassCleanup()
        {
            BaseCleanup();
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        [Test]
        public void SingleCell()
        {
            Assert.That(3d, Is.EqualTo(Worksheet.Cells["B9"].Value));
        }

        [Test]
        public void MultiCell()
        {
            Assert.That(40d, Is.EqualTo(Worksheet.Cells["C9"].Value));
        }

        [Test]
        public void Name()
        {
            Assert.That(10d, Is.EqualTo(Worksheet.Cells["E9"].Value));
        }

        [Test]
        public void ReferenceError()
        {
            Assert.That("#REF!", Is.EqualTo(Worksheet.Cells["H9"].Value.ToString()));
        }

        [Test]
        public void NameOnOtherSheet()
        {
            Assert.That(130d, Is.EqualTo(Worksheet.Cells["I9"].Value));
        }

        [Test]
        public void ArrayInclText()
        {
            Assert.That(7d, Is.EqualTo(Worksheet.Cells["J9"].Value));
        }

        //[Test]
        //public void NameError()
        //{
        //    Assert.That("#NAME?", Is.EqualTo(Worksheet.Cells["L9"].Value));
        //}

        //[Test]
        //public void DivByZeroError()
        //{
        //    Assert.That("#DIV/0!", Is.EqualTo(Worksheet.Cells["M9"].Value));
        //}
    }
}
