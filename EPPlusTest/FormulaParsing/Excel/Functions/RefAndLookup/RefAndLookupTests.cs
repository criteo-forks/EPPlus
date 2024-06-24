using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using FakeItEasy;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class RefAndLookupTests
    {
        const string WorksheetName = null;
        [Test]
        public void LookupArgumentsShouldSetSearchedValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.That(1, Is.EqualTo(lookupArgs.SearchedValue));
        }

        [Test]
        public void LookupArgumentsShouldSetRangeAddress()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.That("A:B", Is.EqualTo(lookupArgs.RangeAddress));
        }

        [Test]
        public void LookupArgumentsShouldSetColIndex()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.That(2, Is.EqualTo(lookupArgs.LookupIndex));
        }

        [Test]
        public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.That(lookupArgs.RangeLookup);
        }

        [Test]
        public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2, true);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.That(lookupArgs.RangeLookup);
        }

        [Test]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns(5);
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100,10));

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns(4);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(1, Is.EqualTo(result.Result));
        }

        [Test]
        public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs("B", "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();;

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns("C");
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(4);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(1, Is.EqualTo(result.Result));
        }

        [Test]
        public void HLookupShouldReturnResultFromMatchingRow()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2, false);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
            Assert.That(expectedResult, Is.EqualTo(result.Result));
        }

        [Test]
        public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(1, "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(result.DataType, Is.EqualTo(DataType.ExcelError));
        }

        [Test]
        public void LookupShouldReturnResultFromMatchingRowArrayVertical()
        {
            var func = new Lookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:B3", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns("B");
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 1)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 2)).Returns("C");
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That("B", Is.EqualTo(result.Result));
        }

        [Test]
        public void LookupShouldReturnResultFromMatchingRowArrayHorizontal()
        {
            var func = new Lookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:C2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns("B");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 3)).Returns("C");

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That("B", Is.EqualTo(result.Result));
        }

        [Test]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontal()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["A3"].Value = "A";
                sheet.Cells["B3"].Value = "B";
                sheet.Cells["C3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, A3:C3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.That("B", Is.EqualTo(result));

            }
        }

        [Test]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontalWithOffset()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["B3"].Value = "A";
                sheet.Cells["C3"].Value = "B";
                sheet.Cells["D3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, B3:D3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.That("B", Is.EqualTo(result));

            } 
        }

        [Test]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(3, "A1:C1", 0);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(3, "A1:A3", 0);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 1)).Returns(5);
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(4, "A1:C1", 1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(6, "A1:C1", -1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(10);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(8);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(10, "A1:C1", -1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(10);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(8);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.That(1, Is.EqualTo(result.Result));
        }

        [Test]
        public void MatchShouldHandleAddressOnOtherSheet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
                sheet2.Cells["A1"].Value = 9;
                sheet2.Cells["A2"].Value = 10;
                sheet2.Cells["A3"].Value = 11;
                sheet1.Calculate();
                Assert.That(2, Is.EqualTo(sheet1.Cells["A1"].Value));
            }    
        }

        [Test]
        public void RowShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
            parsingContext.Scopes.NewScope(rangeAddressFactory.Create("A2"));
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void RowShouldReturnRowSuppliedAddress()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A3"), parsingContext);
            Assert.That(3, Is.EqualTo(result.Result));
        }

        [Test]
        public void ColumnShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
            parsingContext.Scopes.NewScope(rangeAddressFactory.Create("B2"));
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void ColumnShouldReturnRowSuppliedAddress()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("E3"), parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void RowsShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:B3"), parsingContext);
            Assert.That(3, Is.EqualTo(result.Result));
        }

        [Test]
        public void RowsShouldReturnNbrOfRowsForEntireColumn()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A:B"), parsingContext);
            Assert.That(1048576, Is.EqualTo(result.Result));
        }

        [Test]
        public void ColumnssShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Columns();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:E3"), parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void ChooseShouldReturnItemByIndex()
        {
            var func = new Choose();
            var parsingContext = ParsingContext.Create();
            var result = func.Execute(FunctionsHelper.CreateArgs(1, "A", "B"), parsingContext);
            Assert.That("A", Is.EqualTo(result.Result));
        }

        [Test]
        public void AddressShouldReturnAddressByIndexWithDefaultRefType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2), parsingContext);
            Assert.That("$B$1", Is.EqualTo(result.Result));
        }

        [Test]
        public void AddressShouldReturnAddressByIndexWithRelativeType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn), parsingContext);
            Assert.That("B1", Is.EqualTo(result.Result));
        }

        [Test]
        public void AddressShouldReturnAddressByWithSpecifiedWorksheet()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, true, "Worksheet1"), parsingContext);
            Assert.That("Worksheet1!B1", Is.EqualTo(result.Result));
        }

        [Test]
        public void AddressShouldThrowIfR1C1FormatIsSpecified()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                var func = new AddressFunction();
                var parsingContext = ParsingContext.Create();
                parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
                A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
                var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, false), parsingContext); 
            });
        }
    }
}
