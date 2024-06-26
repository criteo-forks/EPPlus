﻿using System;
using System.Collections.Generic;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestFixture]
    public class RowMatcherTests
    {
        private ExcelDatabaseCriteria GetCriteria(Dictionary<ExcelDatabaseCriteriaField, object> items)
        {
            var provider = A.Fake<ExcelDataProvider>();
            var criteria = A.Fake<ExcelDatabaseCriteria>();// (provider, string.Empty);

            A.CallTo(() => criteria.Items).Returns(items);
            return criteria;
        }
        [Test]
        public void IsMatchShouldReturnTrueIfCriteriasMatch()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldReturnFalseIfCriteriasDoesNotMatch()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 4;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(!matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldMatchStrings1()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "1";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldMatchStrings2()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "2";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(!matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldMatchWildcardStrings()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "t*t";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldMatchNumericExpression()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit2")] = "<3";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(matcher.IsMatch(data, criteria));
        }

        [Test]
        public void IsMatchShouldHandleFieldIndex()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField(2)] = "<3";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.That(matcher.IsMatch(data, criteria));
        }
    }
}
