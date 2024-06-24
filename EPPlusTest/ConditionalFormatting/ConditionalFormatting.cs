using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System.IO;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Drawing;

namespace EPPlusTest
{
  /// <summary>
  /// Test the Conditional Formatting feature
  /// </summary>
  [TestFixture]
  public class ConditionalFormatting
  {
    private static ExcelPackage _pck;
    

    #region Additional test attributes
    // You can use the following additional attributes as you write your tests:
    // Use ClassInitialize to run code before running the first test in the class
    [OneTimeSetUp]
    public static void MyClassInitialize()
    {
      if (!Directory.Exists("Test"))
      {
          Directory.CreateDirectory(string.Format("Test"));
      }

      _pck = new ExcelPackage(new FileInfo(@"Test\ConditionalFormatting.xlsx"));
    }

    // Use ClassCleanup to run code after all tests in a class have run
    [OneTimeTearDown]
    public static void MyClassCleanup()
    {
        _pck = null;
    }

    // //Use TestInitialize to run code before running each test 
    // [TestInitialize()]
    // public void MyTestInitialize() 
    // {
    // }

    //// Use TestCleanup to run code after each test has run
    // [TestCleanup()]
    // public void MyTestCleanup() 
    // {
    // }
    #endregion

    /// <summary>
    /// 
    /// </summary>
    [Test]
    public void TwoColorScale()
    {
      var ws = _pck.Workbook.Worksheets.Add("ColorScale");
      ws.ConditionalFormatting.AddTwoColorScale(ws.Cells["A1:A5"]);
      ws.SetValue(1, 1, 1);
      ws.SetValue(2, 1, 2);
      ws.SetValue(3, 1, 3);
      ws.SetValue(4, 1, 4);
      ws.SetValue(5, 1, 5);      
    }

    /// <summary>
    /// 
    /// </summary>
    [Test]
    [Explicit]
    public void ReadConditionalFormatting()
    {
      var pck = new ExcelPackage(new FileInfo(@"c:\temp\cf.xlsx"));

      var ws = pck.Workbook.Worksheets[1];
      Assert.That(ws.ConditionalFormatting.Count == 6);
      Assert.That(ws.ConditionalFormatting[0].Type==eExcelConditionalFormattingRuleType.DataBar);

      var cf1 = ws.ConditionalFormatting.AddEqual(ws.Cells["C3"]);
      //cf1.Formula = "TRUE";
      var cf2 = ws.Cells["C8:C12"].ConditionalFormatting.AddExpression();
      var cf3 = ws.Cells["d12:D22,H12:H22"].ConditionalFormatting.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.RedToBlack);
      pck.SaveAs(new FileInfo(@"c:\temp\cf2.xlsx"));
    }
    /// <summary>
    /// 
    /// </summary>
    [Test]
    [Explicit]
    public void ReadConditionalFormattingError()
    {
      var pck = new ExcelPackage(new FileInfo(@"c:\temp\CofCTemplate.xlsx"));

      var ws = pck.Workbook.Worksheets[1];
      pck.SaveAs(new FileInfo(@"c:\temp\cf2.xlsx"));
    }
    /// <summary>
    /// 
    /// </summary>
    [Test]
    public void TwoBackColor()
    {
        var ws = _pck.Workbook.Worksheets.Add("TwoBackColor");
        IExcelConditionalFormattingEqual condition1 = ws.ConditionalFormatting.AddEqual(ws.Cells["A1"]);
        condition1.StopIfTrue = true;
        condition1.Priority = 1;
        condition1.Formula = "TRUE";
        condition1.Style.Fill.BackgroundColor.Color = Color.Green;
        IExcelConditionalFormattingEqual condition2 = ws.ConditionalFormatting.AddEqual(ws.Cells["A2"]);
        condition2.StopIfTrue = true;
        condition2.Priority = 2;
        condition2.Formula = "FALSE";
        condition2.Style.Fill.BackgroundColor.Color = Color.Red;
    }
    [Test]
    public void Databar()
    {
        var ws = _pck.Workbook.Worksheets.Add("Databar");
        var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
        ws.SetValue(1, 1, 1);
        ws.SetValue(2, 1, 2);
        ws.SetValue(3, 1, 3);
        ws.SetValue(4, 1, 4);
        ws.SetValue(5, 1, 5);
    }
    [Test]   
    public void DatabarChangingAddressAddsConditionalFormatNodeInSchemaOrder()
    {   
        var ws = _pck.Workbook.Worksheets.Add("DatabarAddressing");   
        // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
        ws.HeaderFooter.AlignWithMargins = true;   
        var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);   
        Assert.That("sheetData", Is.EqualTo(cf.Node.ParentNode.PreviousSibling.LocalName));   
        Assert.That("headerFooter", Is.EqualTo(cf.Node.ParentNode.NextSibling.LocalName));   
        cf.Address = new ExcelAddress("C3");   
        Assert.That("sheetData", Is.EqualTo(cf.Node.ParentNode.PreviousSibling.LocalName));   
        Assert.That("headerFooter", Is.EqualTo(cf.Node.ParentNode.NextSibling.LocalName));   
    }
    [Test]
    public void IconSet()
    {
        var ws = _pck.Workbook.Worksheets.Add("IconSet");
        var cf = ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
        ws.SetValue(1, 1, 1);
        ws.SetValue(2, 1, 2);
        ws.SetValue(3, 1, 3);

        var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
        cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
        cf4.Icon1.Formula = "0";
        cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
        cf4.Icon2.Formula = "1/3";
        cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
        cf4.Icon3.Formula = "2/3";
        ws.SetValue(1, 2, 1);
        ws.SetValue(2, 2, 2);
        ws.SetValue(3, 2, 3);
        ws.SetValue(4, 2, 4);

        var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"],eExcelconditionalFormatting5IconsSetType.Quarters);
        cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
        cf5.Icon1.Value = 1;
        cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
        cf5.Icon2.Value = 2;
        cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
        cf5.Icon3.Value = 3;
        cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
        cf5.Icon4.Value = 4;
        cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
        cf5.Icon5.Value = 5;
        cf5.ShowValue = false;
        cf5.Reverse = true;

        ws.SetValue(1, 3, 1);
        ws.SetValue(2, 3, 2);
        ws.SetValue(3, 3, 3);
        ws.SetValue(4, 3, 4);
        ws.SetValue(5, 3, 5);    
    }
    //[Test]
    //public void TwoAndThreeColorConditionalFormattingFromFileDoesNotGetOverwrittenWithDefaultValues()
    //{
    //    var file = new FileInfo(
    //        AppDomain.CurrentDomain.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin"))
    //        + @"Workbooks\MultiColorConditionalFormatting.xlsx");
    //        Assert.That(file.Exists);
    //        using (var package = new ExcelPackage(file))
    //    {
    //        var sheet = package.Workbook.Worksheets.First();
    //        Assert.That(2, Is.EqualTo(sheet.ConditionalFormatting.Count));
    //        var twoColor = (ExcelConditionalFormattingTwoColorScale)sheet.ConditionalFormatting.First(cf => cf is ExcelConditionalFormattingTwoColorScale);
    //        var threeColor = (ExcelConditionalFormattingThreeColorScale)sheet.ConditionalFormatting.First(cf => cf is ExcelConditionalFormattingThreeColorScale);

    //        var defaultTwoColorScale = new ExcelConditionalFormattingTwoColorScale(new ExcelAddress("A1"), 2, sheet);
    //        var defaultThreeColorScale = new ExcelConditionalFormattingThreeColorScale(new ExcelAddress("A1"), 2, sheet);

    //        Assert.That(twoColor.HighValue, Is.Null);
    //        Assert.That(twoColor.LowValue, Is.Null);
    //        Assert.That(defaultTwoColorScale.HighValue, Is.Not.Null);
    //        Assert.That(defaultTwoColorScale.LowValue, Is.Not.Null);
    //        Assert.That(threeColor.HighValue, Is.Null);
    //        Assert.That(threeColor.LowValue, Is.Null);
    //        Assert.That(defaultThreeColorScale.HighValue, Is.Not.Null);
    //        Assert.That(defaultThreeColorScale.LowValue, Is.Not.Null);
    //    }
    //}

    }
}