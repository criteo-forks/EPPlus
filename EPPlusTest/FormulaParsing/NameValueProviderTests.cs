using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class NameValueProviderTests
    {
        //private ExcelDataProvider _excelDataProvider;

        //[SetUp]
        //public void Setup()
        //{
        //    _excelDataProvider = MockRepository.GenerateMock<ExcelDataProvider>();
        //}

        //[Test]
        //public void IsNamedValueShouldReturnTrueIfKeyIsANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("A");
        //    Assert.That(result);
        //}

        //[Test]
        //public void IsNamedValueShouldReturnFalseIfKeyIsNotANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("C");
        //    Assert.That(!result);
        //}

        //[Test]
        //public void GetNamedValueShouldReturnCorrectValueIfKeyExists()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.That("B", Is.EqualTo(result));
        //}

        //[Test]
        //public void ReloadShouldReloadDataFromExcelDataProvider()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.That("B", Is.EqualTo(result));

        //    dict.Clear();
        //    nameValueProvider.Reload();
        //    Assert.That(!nameValueProvider.IsNamedValue("A"));
        //}
    }
}
