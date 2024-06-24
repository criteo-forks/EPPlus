using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class ExcelAddressExpressionTests
    {
        private ParsingContext _parsingContext;
        private ParsingScope _scope;

        private ExcelCell CreateItem(object val)
        {
            return new ExcelCell(val, null, 0, 0);
        }

        [SetUp]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            _scope = _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TearDown]
        public void Cleanup()
        {
            _scope.Dispose();
        }

        [Test]
        public void ConstructorShouldThrowIfExcelDataProviderIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                new ExcelAddressExpression("A1", null, _parsingContext);
            });
            
        }

        [Test]
        public void ConstructorShouldThrowIfParsingContextIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                new ExcelAddressExpression("A1", A.Fake<ExcelDataProvider>(), null);
            });
        }

        //TODO:Fix Test /Janne
        //[Test]
        //public void ShouldCallReturnResultFromProvider()
        //{
        //    var expectedAddress = "A1";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(string.Empty, expectedAddress))
        //        .Return(new object[]{ 1 });

        //    var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
        //    var result = expression.Compile();
        //    Assert.That(1, Is.EqualTo(result.Result));
        //}

        //TODO:Fix Test /Janne
        //[Test]
        //public void CompileShouldReturnAddress()
        //{
        //    var expectedAddress = "A1";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(expectedAddress))
        //        .Return(new ExcelCell[] { CreateItem(1) });

        //    var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
        //    expression.ParentIsLookupFunction = true;
        //    var result = expression.Compile();
        //    Assert.That(expectedAddress, Is.EqualTo(result.Result));

        //}

        #region Compile Tests
        [Test]
        public void CompileSingleCellReference()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    Assert.That(result.Result, Is.Null);
                }
            }
        }

        [Test]
        public void CompileSingleCellReferenceWithValue()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                sheet.Cells[1, 1].Value = "Value";
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    Assert.That("Value", Is.EqualTo(result.Result));
                }
            }
        }

        [Test]
        public void CompileSingleCellReferenceResolveToRange()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("A1", Is.EqualTo(rangeInfo.Address.Address));
                }
            }
        }

        [Test]
        public void CompileSingleCellReferenceResolveToRangeColumnAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("$A1", Is.EqualTo(rangeInfo.Address.Address));
                }
            }
        }

        [Test]
        public void CompileSingleCellReferenceResolveToRangeRowAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("$A1", Is.EqualTo(rangeInfo.Address.Address));
                }
            }
        }

        [Test]
        public void CompileSingleCellReferenceResolveToRangeAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A$1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("$A$1", Is.EqualTo(rangeInfo.Address.Address));
                }
            }
        }

        [Test]
        public void CompileMultiCellReference()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1:A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("A1:A5", Is.EqualTo(rangeInfo.Address.Address));
                    // Enumerating the range still yields no results.
                    Assert.That(0, Is.EqualTo(rangeInfo.Count()));
                }
            }
        }

        [Test]
        public void CompileMultiCellReferenceWithValues()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                sheet.Cells[1, 1].Value = "Value1";
                sheet.Cells[2, 1].Value = "Value2";
                sheet.Cells[3, 1].Value = "Value3";
                sheet.Cells[4, 1].Value = "Value4";
                sheet.Cells[5, 1].Value = "Value5";
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1:A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("A1:A5", Is.EqualTo(rangeInfo.Address.Address));
                    Assert.That(5, Is.EqualTo(rangeInfo.Count()));
                    for (int i = 1; i <= 5; i++)
                    {
                        var rangeItem = rangeInfo.ElementAt(i - 1);
                        Assert.That("Value" + i, Is.EqualTo(rangeItem.Value));
                    }
                }
            }
        }

        [Test]
        public void CompileMultiCellReferenceColumnAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1:$A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("$A1:$A5", Is.EqualTo(rangeInfo.Address.Address));
                    // Enumerating the range still yields no results.
                    Assert.That(0, Is.EqualTo(rangeInfo.Count()));
                }
            }
        }

        [Test]
        public void CompileMultiCellReferenceRowAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A$1:A$5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("A$1:A$5", Is.EqualTo(rangeInfo.Address.Address));
                    // Enumerating the range still yields no results.
                    Assert.That(0, Is.EqualTo(rangeInfo.Count()));
                }
            }
        }

        [Test]
        public void CompileMultiCellReferenceAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A$1:$A$5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.That(rangeInfo, Is.Not.Null);
                    Assert.That("$A$1:$A$5", Is.EqualTo(rangeInfo.Address.Address));
                    // Enumerating the range still yields no results.
                    Assert.That(0, Is.EqualTo(rangeInfo.Count()));
                }
            }
        }
        #endregion
    }
}
