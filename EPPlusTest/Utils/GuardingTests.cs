using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
    [TestFixture]
    public class GuardingTests
    {
        private class TestClass
        {

        }

        [Test]
        public void Require_IsNotNull_ShouldThrowIfArgumentIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                TestClass obj = null;
                Require.Argument(obj).IsNotNull("test");
            });
        }

        [Test]
        public void Require_IsNotNull_ShouldNotThrowIfArgumentIsAnInstance()
        {
            var obj = new TestClass();
            Require.Argument(obj).IsNotNull("test");
        }

        [Test]
        public void Require_IsNotNullOrEmpty_ShouldThrowIfStringIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                string arg = null;
                Require.Argument(arg).IsNotNullOrEmpty("test");
            });
        }

        [Test]
        public void Require_IsNotNullOrEmpty_ShouldNotThrowIfStringIsNotNullOrEmpty()
        {
            string arg = "test";
            Require.Argument(arg).IsNotNullOrEmpty("test");
        }

        [Test]
        public void Require_IsInRange_ShouldThrowIfArgumentIsOutOfRange()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
            {
                int arg = 3;
                Require.Argument(arg).IsInRange(5, 7, "test");
            });
        }

        [Test]
        public void Require_IsInRange_ShouldNotThrowIfArgumentIsInRange()
        {
            int arg = 6;
            Require.Argument(arg).IsInRange(5, 7, "test");
        }
    }
}
