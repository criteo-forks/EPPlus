using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using NUnit.Framework;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;
using System.Text;
using System.Dynamic;
using System.Globalization;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest
{
    /// <summary>
    /// This class contains testcases for issues on Codeplex and Github.
    /// All tests requiering an template should be set to ignored as it's not practical to include all xlsx templates in the project.
    /// </summary>
    [TestFixture]
    public class Issues : TestBase
    {
        [SetUp]
        public void Initialize()
        {
            if (!Directory.Exists(@"c:\Temp"))
            {
                Directory.CreateDirectory(@"c:\Temp");
            }
            if (!Directory.Exists(@"c:\Temp\bug"))
            {
                Directory.CreateDirectory(@"c:\Temp\bug");
            }
        }
        [Test] [Explicit]
        public void Issue15052()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("test");
            ws.Cells["A1:A4"].Value = 1;
            ws.Cells["B1:B4"].Value = 2;
            ws.Cells[1, 1, 4, 1].Style.Numberformat.Format = "#,##0.00;[Red]-#,##0.00";
            ws.Cells[1, 2, 5, 2].Style.Numberformat.Format = "#,##0;[Red]-#,##0";

            p.SaveAs(new FileInfo(@"c:\temp\style.xlsx"));
        }
        [Test]
        public void Issue15041()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = 202100083;
                ws.Cells["A1"].Style.Numberformat.Format = "00\\.00\\.00\\.000\\.0";
                Assert.That("02.02.10.008.3", Is.EqualTo(ws.Cells["A1"].Text));
                ws.Dispose();
            }
        }
        [Test]
        public void Issue15031()
        {
            var d = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(new TimeSpan(35, 59, 1));
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = d;
                ws.Cells["A1"].Style.Numberformat.Format = "[t]:mm:ss";
                ws.Dispose();
            }
        }
        [Test]
        public void Issue15022()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells.AutoFitColumns();
                ws.Cells["A1"].Style.Numberformat.Format = "0";
                ws.Cells.AutoFitColumns();
            }
        }
        [Test]
        public void Issue15056()
        {
            var path = @"C:\temp\output.xlsx";
            var file = new FileInfo(path);
            file.Delete();
            using (var ep = new ExcelPackage(file))
            {
                var s = ep.Workbook.Worksheets.Add("test");
                s.Cells["A1:A2"].Formula = ""; // or null, or non-empty whitespace, with same result
                ep.Save();
            }

        }
        [Explicit]
        [Test]
        public void Issue15058()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\output.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
        }
        [Explicit]
        [Test]
        public void Issue15063()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\TableFormula.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            ws.Calculate();
        }
        [Explicit]
        [Test]
        public void Issue15112()
        {
            System.IO.FileInfo case1 = new System.IO.FileInfo(@"c:\temp\bug\src\src\DeleteRowIssue\Template.xlsx");
            var p = new ExcelPackage(case1);
            var first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case1.xlsx"));

            var case2 = new System.IO.FileInfo(@"c:\temp\bug\src2\DeleteRowIssue\Template.xlsx");
            p = new ExcelPackage(case2);
            first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case2.xlsx"));
        }

        [Explicit]
        [Test]
        public void Issue15118()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bugOutput.xlsx"), new FileInfo(@"c:\temp\bug\DeleteRowIssue\Template.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];
                worksheet.Cells[9, 6, 9, 7].Merge = true;
                worksheet.Cells[9, 8].Merge = false;

                worksheet.DeleteRow(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);

                package.Save();
            }
        }
        [Explicit]
        [Test]
        public void Issue15109()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\test01.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            Assert.That("A1:Z75", Is.EqualTo(ws.Dimension.Address));
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test02.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.That("A1:AF501", Is.EqualTo(ws.Dimension.Address));
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test03.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.That("A1:AD406", Is.EqualTo(ws.Dimension.Address));
            excelP.Dispose();
        }
        [Explicit]
        [Test]
        public void Issue15120()
        {
            var p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\pp.xlsx"));
            ExcelWorksheet ws = p.Workbook.Worksheets["tum_liste"];
            ExcelWorksheet wPvt = p.Workbook.Worksheets.Add("pvtSheet");
            var pvSh = wPvt.PivotTables.Add(wPvt.Cells["B5"], ws.Cells[ws.Dimension.Address.ToString()], "pvtS");

            //p.Save();
        }
        [Test]
        public void Issue15113()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = " Performance Update";
            ws.Cells["A1:H1"].Merge = true;
            ws.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A1:H1"].Style.Font.Size = 14;
            ws.Cells["A1:H1"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["A1:H1"].Style.Font.Bold = true;
            p.SaveAs(new FileInfo(@"c:\temp\merge.xlsx"));
        }
        [Test]
        public void Issue15141()
        {
            using (ExcelPackage package = new ExcelPackage())
            using (ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Test"))
            {
                sheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                sheet.Cells[1, 1, 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                ExcelColumn column = sheet.Column(3); // fails with exception
            }
        }
        [Test] [Explicit]
        public void Issue15145()
        {
            using (ExcelPackage p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\ColumnInsert.xlsx")))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                ws.InsertColumn(12, 3);
                ws.InsertRow(30, 3);
                ws.DeleteRow(31, 1);
                ws.DeleteColumn(7, 1);
                p.SaveAs(new System.IO.FileInfo(@"C:\Temp\bug\InsertCopyFail.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15150()
        {
            var template = new FileInfo(@"c:\temp\bug\ClearIssue.xlsx");
            const string output = @"c:\temp\bug\ClearIssueSave.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[1];
                ws.Cells["A2:C3"].Value = "Test";
                var c = ws.Cells["B2:B3"];
                c.Clear();

                pck.SaveAs(new FileInfo(output));
            }
        }

        [Test] [Explicit]
        public void Issue15146()
        {
            var template = new FileInfo(@"c:\temp\bug\CopyFail.xlsx");
            const string output = @"c:\temp\bug\CopyFail-Save.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[3];

                //ws.InsertColumn(3, 1);
                CustomColumnInsert(ws, 3, 1);

                pck.SaveAs(new FileInfo(output));
            }
        }

        private static void CustomColumnInsert(ExcelWorksheet ws, int column, int columns)
        {
            var source = ws.Cells[1, column, ws.Dimension.End.Row, ws.Dimension.End.Column];
            var dest = ws.Cells[1, column + columns, ws.Dimension.End.Row, ws.Dimension.End.Column + columns];
            source.Copy(dest);
        }
#if !Core
        [Test]
        public void Issue15123()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            using (var dt = new DataTable())
            {
                dt.Columns.Add("String", typeof(string));
                dt.Columns.Add("Int", typeof(int));
                dt.Columns.Add("Bool", typeof(bool));
                dt.Columns.Add("Double", typeof(double));
                dt.Columns.Add("Date", typeof(DateTime));

                var dr = dt.NewRow();
                dr[0] = "Row1";
                dr[1] = 1;
                dr[2] = true;
                dr[3] = 1.5;
                dr[4] = new DateTime(2014, 12, 30);
                dt.Rows.Add(dr);

                dr = dt.NewRow();
                dr[0] = "Row2";
                dr[1] = 2;
                dr[2] = false;
                dr[3] = 2.25;
                dr[4] = new DateTime(2014, 12, 31);
                dt.Rows.Add(dr);

                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells["D2:D3"].Style.Numberformat.Format = "(* #,##0.00);_(* (#,##0.00);_(* \"-\"??_);(@)";

                ws.Cells["E2:E3"].Style.Numberformat.Format = "mm/dd/yyyy";
                ws.Cells.AutoFitColumns();
                Assert.That(ws.Cells[2, 5].Text, Is.Not.EqualTo(""));
            }
        }
#endif
        [Test]
        public void Issue15128()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["B2"].Formula = "A1+$B$1";
            ws.Cells["C1"].Value = "Test";
            ws.Cells["A1:B2"].Copy(ws.Cells["C1"]);
            ws.Cells["B2"].Copy(ws.Cells["D1"]);
            p.SaveAs(new FileInfo(@"c:\temp\bug\copy.xlsx"));
        }

        [Test]
        public void IssueMergedCells()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1:A5,C1:C8"].Merge = true;
            ws.Cells["C1:C8"].Merge = false;
            ws.Cells["A1:A8"].Merge = false;
            p.Dispose();
        }
        [Explicit]
        [Test]
        public void Issue15158()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\Output.xlsx"), new FileInfo(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet worksheet = workBook.Worksheets[1];

                //string column = ColumnIndexToColumnLetter(28);
                worksheet.DeleteColumn(28);

                if (worksheet.Cells["AA19"].Formula != "")
                {
                    throw new Exception("this cell should not have formula");
                }

                package.Save();
            }
        }

        public class cls1
        {
            public int prop1 { get; set; }
        }

        public class cls2 : cls1
        {
            public string prop2 { get; set; }
        }
        [Test]
        public void LoadFromColIssue()
        {
            var l = new List<cls1>();

            l.Add(new cls1() { prop1 = 1 });
            l.Add(new cls2() { prop1 = 1, prop2 = "test1" });

            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Test");

            ws.Cells["A1"].LoadFromCollection(l, true, TableStyles.Light16, BindingFlags.Instance | BindingFlags.Public,
                new MemberInfo[] { typeof(cls2).GetProperty("prop2") });
        }

        [Test]
        public void Issue15168()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Cells[1, 1].Value = "A1";
                ws.Cells[2, 1].Value = "A2";

                ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
                Assert.Equals("A1", ws.Cells[1, 1].Value);
            }
        }
        [Explicit]
        [Test]
        public void Issue15159()
        {
            var fs = new FileStream(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx", FileMode.OpenOrCreate);
            using (var package = new OfficeOpenXml.ExcelPackage(fs))
            {
                package.Save();
            }
            fs.Seek(0, SeekOrigin.Begin);
            var fs2 = fs;
        }
        [Test]
        public void Issue15179()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
                ws.Cells["E3:F3"].Merge = true;
                ws.Cells["E3:F3"].Merge = false;
                ws.DeleteRow(2, 6);
                ws.Cells["A1"].Value = 0;
                var s = ws.Cells["A1"].Value.ToString();

            }
        }
        [Explicit]
        [Test]
        public void Issue15169()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\issue\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Labour Costs";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"C:\temp\bug\issue\output2.xlsx"));
            }
        }
        [Explicit]
        [Test]
        public void Issue15172()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\book2.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                Assert.Equals("IF($R10>=X$2,1,0)", ws.Cells["X10"].Formula);
                ws.Calculate();
                Assert.That(0D, Is.EqualTo(ws.Cells["X10"].Value));
            }
        }
        [Explicit]
        [Test]
        public void Issue15174()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\temp\bug\MyTemplate.xlsx")))
            {
                package.Workbook.Worksheets[1].Column(2).Style.Numberformat.Format = "dd/mm/yyyy";

                package.SaveAs(new FileInfo(@"C:\temp\bug\MyTemplate2.xlsx"));
            }
        }
        [Explicit]
        [Test]
        public void PictureIssue()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Drawings.AddPicture("Test", new FileInfo(@"c:\temp\bug\2152228.jpg"));
            p.SaveAs(new FileInfo(@"c:\temp\bug\pic.xlsx"));
        }

        [Explicit]
        [Test]
        public void Issue14988()
        {
            var guid = Guid.NewGuid().ToString("N");
            using (var outputStream = new FileStream(@"C:\temp\" + guid + ".xlsx", FileMode.Create))
            {
                using (var inputStream = new FileStream(@"C:\temp\bug2.xlsx", FileMode.Open))
                {
                    using (var package = new ExcelPackage(outputStream, inputStream, "Test"))
                    {
                        var ws = package.Workbook.Worksheets.Add("Test empty");
                        ws.Cells["A1"].Value = "Test";
                        package.Encryption.Password = "Test2";
                        package.Save();
                        //package.SaveAs(new FileInfo(@"c:\temp\test2.xlsx"));
                    }
                }
            }
        }

        [Test] [Explicit]
        public void Issue15173_1()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel01.xlsx")))
            {
                var sw = new Stopwatch();
                //pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                pck.Workbook.Calculate();
                Assert.That("20L2300", Is.EqualTo(ws.Cells["F4"].Value));
                Assert.That("20K2E01", Is.EqualTo(ws.Cells["F5"].Value));
                var f7Val = pck.Workbook.Worksheets["MODELLO-TIPO PANNELLO"].Cells["F7"].Value;
                Assert.Equals(13.445419, Math.Round((double)f7Val, 6));
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }

        [Test] [Explicit]
        public void Issue15173_2()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel02.xlsx")))
            {
                var sw = new Stopwatch();
                pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                //ws.Calculate();
                pck.Workbook.Calculate();
                Assert.That("20L2300", Is.EqualTo(ws.Cells["F4"].Value));
                Assert.That("20K2E01", Is.EqualTo(ws.Cells["F5"].Value));
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }
        [Explicit]
        [Test]
        public void Issue15154()
        {
            Directory.EnumerateFiles(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\").AsParallel().ForAll(file =>
            {
                //lock (_lock)
                //{
                using (var package = new ExcelPackage(new FileStream(file, FileMode.Open)))
                {
                    package.Workbook.Worksheets[1].Cells[1, 1].Value = file;
                    package.SaveAs(new FileInfo(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\new\" + new FileInfo(file).Name));
                }
                //}
            });

        }
        [Explicit]
        [Test]
        public void Issue15188()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("test");
                worksheet.Column(6).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(7).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(8).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(10).Style.Numberformat.Format = "mm/dd/yyyy";

                worksheet.Cells[2, 6].Value = DateTime.Today;
                string a = worksheet.Cells[2, 6].Text;
                Assert.That(DateTime.Today.ToString("MM/dd/yyyy"), Is.EqualTo(a));
            }
        }
        [Test] [Explicit]
        public void Issue15194()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15194-Save.xlsx"), new FileInfo(@"c:\temp\bug\I15194.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.Cells["E3:F3"].Merge = false;

                worksheet.DeleteRow(2, 6);

                package.Save();
            }
        }
        [Test] [Explicit]
        public void Issue15195()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\i15195.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }
        [Test] [Explicit]
        public void Issue14788()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\GetWorkSheetXmlBad.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }
        [Test] [Explicit]
        public void Issue15167()
        {
            FileInfo fileInfo = new FileInfo(@"c:\temp\bug\Draw\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Board pack";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"c:\temp\bug\output.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15198()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Output.xlsx"), new FileInfo(@"c:\temp\bug\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.DeleteRow(12);

                package.Save();
            }
        }
        [Test] [Explicit]
        public void Issue13492()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Bug13492.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                var rt = worksheet.Cells["K31"].RichText.Text;

                package.Save();
            }
        }
        [Test] [Explicit]
        public void Issue14966()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ssis\FileFromReportingServer2012.xlsx")))
                package.SaveAs(new FileInfo(@"c:\temp\bug\ssis\Corrupted.xlsx"));
        }
        [Test] [Explicit]
        public void Issue15200()
        {
            File.Copy(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\input.xlsx", @"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx", true);

            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells[1, 1, 1, 2];
                var resultRange = sheet.Cells[3, 1, 3, 2];
                sourceRange.Copy(resultRange);

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[5, 1, 5, 7];
                sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[7, 3, 7, 7];
                sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

                p.Save();
            }
        }
        [Test]
        public void Issue15212()
        {
            var s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleBug");
                ws.Cells["A1"].Value = 5698633.64;
                ws.Cells["A1"].Style.Numberformat.Format = s;
                var t = ws.Cells["A1"].Text;
            }
        }
        [Test] [Explicit]
        public void Issue15213()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile.xlsx")))
            {
                foreach (var ws in p.Workbook.Worksheets)
                {
                    ws.Cells[1023, 1, ws.Dimension.End.Row - 2, ws.Dimension.End.Column].Clear();
                    Assert.That(ws.Dimension, Is.Not.Null);
                }
                foreach (var cell in p.Workbook.Worksheets[2].Cells)
                {
                    Console.WriteLine(cell);
                }
                p.SaveAs(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile-save.xlsx"));
            }
        }

        [Test] [Explicit]
        public void Issuer15217()
        {

            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\FormatRowCol.xlsx")))
            {
                var ws = p.Workbook.Worksheets.Add("fmt");
                ws.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Row(1).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells["A1:B2"].Value = 1;
                ws.Column(1).Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
                ws.Column(2).Style.Numberformat.Format = "#,##0";
                p.Save();
            }
        }
        [Test] [Explicit]
        public void Issuer15228()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("colBug");

                var col = ws.Column(7);
                col.ColumnMax = 8;
                col.Hidden = true;

                var col8 = ws.Column(8);
                Assert.That(true, Is.EqualTo(col8.Hidden));
            }
        }

        [Test] [Explicit]
        public void Issue15234()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\merge2\input.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells["1:4"];

                sheet.InsertRow(5, 4);

                var resultRange = sheet.Cells["5:8"];
                sourceRange.Copy(resultRange);

                p.Save();
            }
        }
        [Test]
        /**** Pivottable issue ****/
        public void Issue()
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"c:\ExcelPivotTest");
            FileInfo MyFile = new FileInfo(@"c:\temp\bug\pivottable.xlsx");
            LoadData(MyFile);
            BuildPivotTable1(MyFile);
            BuildPivotTable2(MyFile);
        }

        private void LoadData(FileInfo MyFile)
        {
            if (MyFile.Exists)
            {
                MyFile.Delete();  // ensures we create a new workbook
            }

            using (ExcelPackage EP = new ExcelPackage(MyFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet wsData = EP.Workbook.Worksheets.Add("Data");
                //Add the headers
                wsData.Cells[1, 1].Value = "INVOICE_DATE";
                wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
                wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
                wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
                wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
                wsData.Cells[1, 6].Value = "COUNT";

                //Add some items...
                wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B2"].Value = 33.63;
                wsData.Cells["C2"].Value = (-.87);
                wsData.Cells["D2"].Value = "Unfavorable Price Variance";
                wsData.Cells["E2"].Value = "Pending";
                wsData.Cells["F2"].Value = 1;

                wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B3"].Value = 43.14;
                wsData.Cells["C3"].Value = (-1.29);
                wsData.Cells["D3"].Value = "Unfavorable Price Variance";
                wsData.Cells["E3"].Value = "Pending";
                wsData.Cells["F3"].Value = 1;

                wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B4"].Value = 55;
                wsData.Cells["C4"].Value = (-2.87);
                wsData.Cells["D4"].Value = "Unfavorable Price Variance";
                wsData.Cells["E4"].Value = "Pending";
                wsData.Cells["F4"].Value = 1;

                wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B5"].Value = 38.72;
                wsData.Cells["C5"].Value = (-5.00);
                wsData.Cells["D5"].Value = "Unfavorable Price Variance";
                wsData.Cells["E5"].Value = "Pending";
                wsData.Cells["F5"].Value = 1;

                wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B6"].Value = 77.44;
                wsData.Cells["C6"].Value = (-1.55);
                wsData.Cells["D6"].Value = "Unfavorable Price Variance";
                wsData.Cells["E6"].Value = "Pending";
                wsData.Cells["F6"].Value = 1;

                wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B7"].Value = 127.55;
                wsData.Cells["C7"].Value = (-10.50);
                wsData.Cells["D7"].Value = "Unfavorable Price Variance";
                wsData.Cells["E7"].Value = "Pending";
                wsData.Cells["F7"].Value = 1;

                using (var range = wsData.Cells[2, 1, 7, 1])
                {
                    range.Style.Numberformat.Format = "mm-dd-yy";
                }

                wsData.Cells.AutoFitColumns(0);
                EP.Save();
            }
        }
        private void BuildPivotTable1(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot1");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
                pivotTable1.ColumnGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }

        private void BuildPivotTable2(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot2");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
                pivotTable1.ColumnGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }

        [Test] [Explicit]
        public void issue15249()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Boldtextcopy.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\Boldtextcopy2.xlsx"));
            }
        }
        [Test] [Explicit]
        public void issue15300()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\headfootpic.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\headfootpic_save.xlsx"));
            }

        }
        [Test] [Explicit]
        public void issue15295()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivot issue\input.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivotcoldup.xlsx"));
            }

        }
        [Test] [Explicit]
        public void issue15282()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivottable-table.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivottab-tab-save.xlsx"));
            }

        }

        [Test] [Explicit]
        public void Issues14699()
        {
            FileInfo newFile = new FileInfo(string.Format("c:\\temp\\bug\\EPPlus_Issue14699.xlsx", System.IO.Directory.GetCurrentDirectory()));
            OfficeOpenXml.ExcelPackage pkg = new ExcelPackage(newFile);
            ExcelWorksheet wksheet = pkg.Workbook.Worksheets.Add("Issue14699");
            // Initialize a small range
            for (int row = 1; row < 11; row++)
            {
                for (int col = 1; col < 11; col++)
                {
                    wksheet.Cells[row, col].Value = string.Format("{0}{1}", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col - 1], row);
                }
            }
            wksheet.View.FreezePanes(3, 3);
            pkg.Save();

        }
        [Test] [Explicit]
        public void Issue15382()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Text Run Issue.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\inlinText.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15380()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\dotinname.xlsx")))
            {
                var v = exfile.Workbook.Worksheets["sheet1.3"].Names["Test.Name"].Value;
                Assert.That(v, Is.EqualTo(1));
            }
        }
        [Test] [Explicit]
        public void Issue15378()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bubble.xlsx")))
            {
                var c = p.Workbook.Worksheets[1].Drawings[0] as ExcelBubbleChart;
                var cs = c.Series[0] as ExcelBubbleChartSerie;
            }
        }
        [Test]
        public void Issue15377()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws1");
                ws.Cells["A1"].Value = (double?)1;
                var v = ws.GetValue<double?>(1, 1);
            }
        }
        [Test]
        public void Issue15374()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RT");
                var r = ws.Cells["A1"];
                r.RichText.Text = "Cell 1";
                r["A2"].RichText.Add("Cell 2");
                p.SaveAs(new FileInfo(@"c:\temp\rt.xlsx"));
            }
        }
        [Test]
        public void IssueTranslate()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Trans");
                ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
                var fr = ws.Cells["A1:A2"].FormulaR1C1;
                ws.Cells["A1:A2"].FormulaR1C1 = fr;
                Assert.Equals("IF(1=1,\"A's B C\",\"D\")", ws.Cells["A2"].Formula);
            }
        }
        [Test]
        public void Issue15397()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("styleerror");
                workSheet.Cells["F:G"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                //And then: 

                workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

                workSheet.Cells["I:I"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

                workSheet.InsertColumn(2, 2, 9);
                workSheet.Column(45).Width = 0;

                p.SaveAs(new FileInfo(@"c:\temp\styleerror.xlsx"));
            }
        }
        [Test]
        public void Issuer14801()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("rterror");
                var cell = workSheet.Cells["A1"];
                cell.RichText.Add("toto: ");
                cell.RichText[0].PreserveSpace = true;
                cell.RichText[0].Bold = true;
                cell.RichText.Add("tata");
                cell.RichText[1].Bold = false;
                cell.RichText[1].Color = Color.Green;
                p.SaveAs(new FileInfo(@"c:\temp\rtpreserve.xlsx"));
            }
        }
        [Test]
        public void Issuer15445()
        {
            using (var p = new ExcelPackage())
            {
                var ws1 = p.Workbook.Worksheets.Add("ws1");
                var ws2 = p.Workbook.Worksheets.Add("ws2");
                ws2.View.SelectedRange = "A1:B3 D12:D15";
                ws2.View.ActiveCell = "D15";
                p.SaveAs(new FileInfo(@"c:\temp\activeCell.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15429()
        {
            FileInfo file = new FileInfo(@"c:\temp\original.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                var equalsRule = worksheet.ConditionalFormatting.AddEqual(new ExcelAddress(2, 3, 6, 3));
                equalsRule.Formula = "0";
                equalsRule.Style.Fill.BackgroundColor.Color = Color.Blue;
                worksheet.ConditionalFormatting.AddDatabar(new ExcelAddress(4, 4, 4, 4), Color.Red);
                excelPackage.Save();
            }
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets["Sheet 1"];
                int i = 0;
                foreach (var conditionalFormat in worksheet.ConditionalFormatting)
                {
                    conditionalFormat.Address = new ExcelAddress(5 + i++, 5, 6, 6);
                }
                excelPackage.SaveAs(new FileInfo(@"c:\temp\error.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15436()
        {
            FileInfo file = new FileInfo(@"c:\temp\incorrect value.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                Assert.That(excelPackage.Workbook.Worksheets[1].Cells["A1"].Value, Is.EqualTo(19120072));
            }
        }
        [Test] [Explicit]
        public void Issue13128()
        {
            FileInfo file = new FileInfo(@"c:\temp\students.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                Assert.That(((ExcelChart)excelPackage.Workbook.Worksheets[1].Drawings[0]).Series[0].XSeries, Is.Not.Null);
            }
        }
        [Test] [Explicit]
        public void Issue15252()
        {
            using (var p = new ExcelPackage())
            {
                var path1 = @"c:\temp\saveerror1.xlsx";
                var path2 = @"c:\temp\saveerror2.xlsx";
                var workSheet = p.Workbook.Worksheets.Add("saveerror");
                workSheet.Cells["A1"].Value = "test";

                // double save OK?
                p.SaveAs(new FileInfo(path1));
                p.SaveAs(new FileInfo(path2));

                // files are identical?
#if (Core)
                var md5 = System.Security.Cryptography.MD5.Create();
#else
                var md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
#endif
                using (var fs1 = new FileStream(path1, FileMode.Open))
                using (var fs2 = new FileStream(path2, FileMode.Open))
                {
                    var hash1 = String.Join("", md5.ComputeHash(fs1).Select((x) => { return x.ToString(); }));
                    var hash2 = String.Join("", md5.ComputeHash(fs2).Select((x) => { return x.ToString(); }));
                    Assert.That(hash1, Is.EqualTo(hash2));
                }
            }
        }
        [Test] [Explicit]
        public void Issue15469()
        {
            ExcelPackage excelPackage = new ExcelPackage(new FileInfo(@"c:\temp\bug\EPPlus-Bug.xlsx"), true);
            using (FileStream fs = new FileStream(@"c:\temp\bug\EPPlus-Bug-new.xlsx", FileMode.Create))
            {
                excelPackage.SaveAs(fs);
            }
        }
        [Test]
        public void Issue15438()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                var c = ws.Cells["A1"].Style.Font.Color;
                c.Indexed = 3;
                Assert.That(c.LookupColor(c), Is.EqualTo("#FF00FF00"));
            }
        }
        [Test] [Explicit]
        public void Issue15097()
        {
            using (var pkg = new ExcelPackage())
            {
                var templateFile = ReadTemplateFile(@"c:\temp\bug\test_vorlage3.xlsx");
                using (var ms = new System.IO.MemoryStream(templateFile))
                {
                    using (var tempPkg = new ExcelPackage(ms))
                    {
                        tempPkg.Workbook.Worksheets.Copy(tempPkg.Workbook.Worksheets.First().Name, "Demo");
                    }
                }
            }
        }
        [Test] [Explicit]
        public void Issue15485()
        {
            using (var pkg = new ExcelPackage(new FileInfo(@"c:\temp\bug\PivotChartSeriesIssue.xlsx")))
            {
                var ws = pkg.Workbook.Worksheets[1];
                ws.InsertRow(1, 1);
                ws.InsertColumn(1, 1);
                pkg.Save();
            }
        }
        public static byte[] ReadTemplateFile(string templateName)
        {
            byte[] templateFIle;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (var sw = new System.IO.FileStream(templateName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    byte[] buffer = new byte[2048];
                    int bytesRead;
                    while ((bytesRead = sw.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, bytesRead);
                    }
                }
                ms.Position = 0;
                templateFIle = ms.ToArray();
            }
            return templateFIle;

        }

        [Test]
        public void Issue15455()
        {
            using (var pck = new ExcelPackage())
            {

                var sheet1 = pck.Workbook.Worksheets.Add("sheet1");
                var sheet2 = pck.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["C2"].Value = 3;
                sheet1.Cells["C3"].Formula = "VLOOKUP(E1, Sheet2!A1:D6, C2, 0)";
                sheet1.Cells["E1"].Value = "d";

                sheet2.Cells["A1"].Value = "d";
                sheet2.Cells["C1"].Value = "dg";
                pck.Workbook.Calculate();
                var c3 = sheet1.Cells["C3"].Value;
                Assert.That("dg", Is.EqualTo(c3));
            }
        }

        [Test]
        public void Issue15460WithString()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] { "value1", "value2", "value3" };
                    package.Save();
                }
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.Equals("value1", sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }

        [Test]
        public void Issue15460WithNull()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] { null, "value2", "value3" };
                    package.Save();
                }
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.Equals(string.Empty, sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }

        [Test]
        public void Issue15460WithNonStringPrimitive()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] { 5, 6, 7 };
                    package.Save();
                }
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.Equals((double)5, sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }
        [Test] [Explicit]
        public void MergeIssue()
        {
            var worksheetPath = Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");
            FileInfo fi = new FileInfo(Path.Combine(worksheetPath, "Example.xlsx"));
            fi.Delete();
            using (ExcelPackage pckg = new ExcelPackage(fi))
            {
                var ws = pckg.Workbook.Worksheets.Add("Example");
                ws.Cells[1, 1, 1, 3].Merge = true;
                ws.Cells[1, 1, 1, 3].Merge = true;
                pckg.Save();
            }
        }
        [Test] [Explicit]
        public void Issuer15563()   //And 15562
        {
            using (var package = new ExcelPackage())
            {
                var w = package.Workbook.Worksheets.Add("test");
                w.Row(1).Style.Font.Bold = true;
                w.Row(2).Style.Font.Bold = true;

                for (var i = 0; i < 4; i++)
                {
                    w.Column(8 + 2 * i).Style.Border.Right.Style = ExcelBorderStyle.Dotted;
                }
                package.SaveAs(new FileInfo(@"c:\temp\bug\stylebug.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issuer15560()
        {
            //Type not set to error when converting shared formula.
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\sharedFormulas.xlsm")))
            {
                package.SaveAs(new FileInfo(@"c:\temp\bug\sharedformulabug.xlsm"));
            }
        }
        [Test] [Explicit]
        public void Issuer15558()
        {
            //TODO: ??? works
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\test_file_20161118.xlsx")))
            {
                package.SaveAs(new FileInfo(@"c:\temp\bug\saveproblem.xlsm"));
            }
        }
        /// <summary>
        /// Issue 15561
        /// </summary>
        [Test] [Explicit]
        public void Chart_From_Cell_Union_Selector_Bug_Test()
        {
            var existingFile = new FileInfo(@"c:\temp\Chart_From_Cell_Union_Selector_Bug_Test.xlsx");
            if (existingFile.Exists)
                existingFile.Delete();

            using (var pck = new ExcelPackage(existingFile))
            {
                var myWorkSheet = pck.Workbook.Worksheets.Add("Content");
                var ExcelWorksheet = pck.Workbook.Worksheets.Add("Chart");

                //Some data
                myWorkSheet.Cells["A1"].Value = "A";
                myWorkSheet.Cells["A2"].Value = 100; myWorkSheet.Cells["A3"].Value = 400; myWorkSheet.Cells["A4"].Value = 200; myWorkSheet.Cells["A5"].Value = 300; myWorkSheet.Cells["A6"].Value = 600; myWorkSheet.Cells["A7"].Value = 500;
                myWorkSheet.Cells["B1"].Value = "B";
                myWorkSheet.Cells["B2"].Value = 300; myWorkSheet.Cells["B3"].Value = 200; myWorkSheet.Cells["B4"].Value = 1000; myWorkSheet.Cells["B5"].Value = 600; myWorkSheet.Cells["B6"].Value = 500; myWorkSheet.Cells["B7"].Value = 200;

                //Pie chart shows with EXTRA B2 entry due to problem with ExcelRange Enumerator
                ExcelRange values = myWorkSheet.Cells["B2,B4,B6"];  //when the iterator is evaluated it will return the first cell twice: "B2,B2,B4,B6"
                ExcelRange xvalues = myWorkSheet.Cells["A2,A4,A6"]; //when the iterator is evaluated it will return the first cell twice: "A2,A2,A4,A6"
                var chartBug = ExcelWorksheet.Drawings.AddChart("Chart BAD", eChartType.Pie);
                chartBug.Series.Add(values, xvalues);
                chartBug.Title.Text = "Using ExcelRange";

                //Pie chart shows correctly when using string addresses and avoiding ExcelRange
                var chartGood = ExcelWorksheet.Drawings.AddChart("Chart GOOD", eChartType.Pie);
                chartGood.SetPosition(10, 0, 0, 0);
                chartGood.Series.Add("Content!B2,Content!B4,Content!B6", "Content!A2,Content!A4,Content!A6");
                chartGood.Title.Text = "Using String References";

                pck.Save();
            }
        }
        [Test] [Explicit]
        public void Issue15566()
        {
            string TemplateFileName = @"c:\temp\bug\TestWithPivotTablePointingToExcelTableForData.xlsx";
            string ExportFileName = @"c:\temp\bug\TestWithPivotTablePointingToExcelTableForData_Export.xlsx";
            using (ExcelPackage excelpackage = new ExcelPackage(new FileInfo(TemplateFileName), true))
            {
                excelpackage.SaveAs(new FileInfo(ExportFileName));
            }
        }
        [Test] [Explicit]
        public void Issue15564()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FormulaBug");
                for (int i = 1; i < 1030; i++)
                {
                    ws.Cells[i, 1].Value = i;
                    ws.Cells[i, 2].FormulaR1C1 = "rc[-1]+1";
                }
                ws.InsertRow(4, 1025, 3);
                ws.InsertRow(1050, 1025, 3);
                p.SaveAs(new FileInfo(@"c:\temp\bug\fb.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue15551()    //Works fine?
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleBug");

                ws.Cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                p.SaveAs(new FileInfo(@"c:\temp\bug\StyleBug.xlsx"));
            }
        }
        [Test] [Explicit]
        public void BugCommentNullAfterRemove()
        {

            string xls = @"c:\temp\bug\in.xlsx";

            ExcelPackage theExcel = new ExcelPackage(new FileInfo(xls), true);

            ExcelRangeBase cell;
            ExcelComment cmnt;
            ExcelWorksheet ws = theExcel.Workbook.Worksheets[1];

            foreach (string addr in "B4 B11".Split(' '))
            {
                cell = ws.Cells[addr];
                cmnt = cell.Comment;

                Assert.That(cmnt, Is.Not.Null, "Comment in " + addr + " expected not null ");
                ws.Comments.Remove(cmnt);
            }

        }

        [Test] [Explicit]
        public void BugCommentExceptionOnRemove()
        {

            string xls = @"c:\temp\bug\in.xlsx";

            ExcelPackage theExcel = new ExcelPackage(new FileInfo(xls), true);

            ExcelRangeBase cell;
            ExcelComment cmnt;
            ExcelWorksheet ws = theExcel.Workbook.Worksheets[1];

            foreach (string addr in "B4 B16".Split(' '))
            {
                cell = ws.Cells[addr];
                cmnt = cell.Comment;

                try
                {
                    ws.Comments.Remove(cmnt);
                }
                catch (Exception ex)
                {
                    Assert.Fail("Exception while removing comment at " + addr + ": " + ex.Message);
                }
            }

        }

        [Test]
        public void Issue15548_SumIfsShouldHandleGaps()
        {
            using (var package = new ExcelPackage())
            {
                var test = package.Workbook.Worksheets.Add("Test");

                test.Cells["A1"].Value = 1;
                test.Cells["B1"].Value = "A";

                //test.Cells["A2"] is default
                test.Cells["B2"].Value = "A";

                test.Cells["A3"].Value = 1;
                test.Cells["B4"].Value = "B";

                test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

                test.Calculate();

                var result = test.Cells["D2"].GetValue<int>();

                Assert.That(1, Is.EqualTo(result), string.Format("Expected 1, got {0}", result));
            }
        }
        [Test]
        public void Issue15548_SumIfsShouldHandleBadData()
        {
            using (var package = new ExcelPackage())
            {
                var test = package.Workbook.Worksheets.Add("Test");

                test.Cells["A1"].Value = 1;
                test.Cells["B1"].Value = "A";

                test.Cells["A2"].Value = "Not a number";
                test.Cells["B2"].Value = "A";

                test.Cells["A3"].Value = 1;
                test.Cells["B4"].Value = "B";

                test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

                test.Calculate();

                var result = test.Cells["D2"].GetValue<int>();

                Assert.That(1, Is.EqualTo(result), string.Format("Expected 1, got {0}", result));
            }
        }
        [Test] [Explicit]
        public void Issue_15585()
        {
            var excelFile = new FileInfo(@"c:\temp\bug\formula_value.xlsx");
            using (var package = new ExcelPackage(excelFile))
            {
                // Output from the logger will be written to the following file
                var logfile = new FileInfo(@"C:\temp\EpplusLogFile.txt");
                // Attach the logger before the calculation is performed.
                package.Workbook.FormulaParserManager.AttachLogger(logfile);
                // Calculate - can also be executed on sheet- or range level.
                package.Workbook.Calculate();

                Console.WriteLine(String.Format("Country: \t\t\t{0}", package.Workbook.Worksheets[1].Cells["B1"].Value));
                Console.WriteLine(String.Format("Phone Code - Direct Reference:\t{0}", package.Workbook.Worksheets[1].Cells["B2"].Value.ToString()));
                Console.WriteLine(String.Format("Phone Code - Name Ranges:\t{0}", package.Workbook.Worksheets[1].Cells["B3"].Value.ToString()));
                Console.WriteLine(String.Format("Phone Code - Table reference:\t{0}", package.Workbook.Worksheets[1].Cells["B4"].Value.ToString()));

                // The following method removes any logger attached to the workbook.
                package.Workbook.FormulaParserManager.DetachLogger();
            }
        }
        [Test]
        public void Issue_15641()
        {
            ExcelPackage ep = new ExcelPackage();

            ExcelWorksheet sheet = ep.Workbook.Worksheets.Add("A Sheet");

            sheet.Cells[1, 1].CreateArrayFormula("IF(SUM(B1:J1)>0,SUM(B2:J2))"); //A1
            sheet.Cells[2, 1].Value = sheet.Cells[1, 1].IsArrayFormula; //A2
            sheet.Cells[1, 1].Copy(sheet.Cells[3, 1]); //A3

            Assert.Equals(true, sheet.Cells[3, 1].IsArrayFormula);
        }
        [Test] [Explicit]
        public void Issue_5()
        {
            var excelFile = new FileInfo(@"c:\temp\bug\test.xlsm");
            using (var package = new ExcelPackage(excelFile))
            {
                var ws = package.Workbook.Worksheets.Add("NewWorksheet");
                ws.CodeModule.Code = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n\r\nEnd Sub";
                package.SaveAs(new FileInfo(@"c:\temp\bug\vbafailSaved.xlsm"));
            }
        }
        [Test] [Explicit]
        public void Issue_8()
        {
            dynamic c = 1;

            var l = new List<dynamic>();
            l.Add(1);
            l.Add("s");
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Dynamic Test");
                ws.Cells["a1"].LoadFromCollection(l);
                package.SaveAs(new FileInfo(@"c:\temp\dynamic.xlsx"));
            }


        }
        [Test] [Explicit]
        public void Issuer27()
        {
            FileInfo file = new FileInfo(@"C:\Temp\Test.xlsx");
            var pck = new ExcelPackage(file);
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Worksheet");
            if (ws != null)
            {
                ws.Cells["A1"].Value = "Cell value 1";
                ws.Cells["B1"].Value = "Cell value 2";
                ws.Cells["C1"].Value = "Cell value 3";
                ws.Cells["D1"].Value = "Cell value 4";
                ws.Cells["E1"].Value = "Cell value 5";
            }
            //ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment::Top; // Columns 4 and greater hidden
            ws.Cells.AutoFitColumns(0);
            ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top; // 4.1.1.0 - exception, 4.0.4.0 - columns 4 and 5 hidden
            ws.Column(4).Hidden = true;
            ws.Column(5).Hidden = true; // span exception
            pck.Save();
        }
        [Test] [Explicit]
        public void Issuer26()
        {
            FileInfo file = new FileInfo(@"C:\Temp\repeatrowcol.xlsx");
            var pck = new ExcelPackage(file);
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Worksheet");
            ws.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
            ws.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");
            pck.Save();
        }
        [Test] [Explicit]
        public void Issue32()
        {
            var outputDir = new DirectoryInfo(@"c:\temp\sampleapp");
            var existFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            string newFileName = outputDir.FullName + @"\sample1_copied.xlsx";

            System.IO.File.Copy(existFile.FullName, newFileName, true);
            var newFile = new FileInfo(newFileName);

            using (var package = new ExcelPackage(newFile))
            {
                // Add a new worksheet to the empty workbook

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory_copied", package.Workbook.Worksheets[1]);

                //ExcelNamedRange sourceRange = package.Workbook.Names["Body"];
                ExcelNamedRange sourceRange = package.Workbook.Names["row"];
                ExcelNamedRange sourceRange2 = package.Workbook.Names["roww"];
                ExcelNamedRange sourceRange3 = package.Workbook.Names["rowww"];
                ExcelWorksheet worksheetFrom = sourceRange.Worksheet;

                ExcelWorksheet worksheetTo = package.Workbook.Worksheets["Inventory_copied"];

                ExcelRange cells = worksheetTo.Cells;
                ExcelRange rangeTo = cells[1, 1, 1, 16384];
                sourceRange.Copy(rangeTo);
                ExcelRange rangeTo2 = cells[2, 1, 2, 16384];
                sourceRange2.Copy(rangeTo2);
                ExcelRange rangeTo3 = cells[4, 1, 4, 16384];
                sourceRange3.Copy(rangeTo3);

                package.Save();

                //}

            }
        }
        [Test]
        public void Issue63() // See https://github.com/JanKallman/EPPlus/issues/63
        {
            // Prepare
            var newFile = new FileInfo(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx"));
            try
            {
                using (var package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("ArrayTest");
                    ws.Cells["A1"].Value = 1;
                    ws.Cells["A2"].Value = 2;
                    ws.Cells["A3"].Value = 3;
                    ws.Cells["B1:B3"].CreateArrayFormula("A1:A3");
                    package.Save();
                }
                Assert.That(File.Exists(newFile.FullName));

                // Test: basic support to recognize array formulas after reading Excel workbook file
                using (var package = new ExcelPackage(newFile))
                {
                    Assert.That("A1:A3", Is.EqualTo(package.Workbook.Worksheets["ArrayTest"].Cells["B1"].Formula));
                    Assert.That(package.Workbook.Worksheets["ArrayTest"].Cells["B1"].IsArrayFormula);
                }
            }
            finally
            {
                File.Delete(newFile.FullName);
            }
        }
        [Test] [Explicit]
        public void Issue60()
        {
            var ms = new MemoryStream(File.ReadAllBytes(@"c:\temp\sampleapp\sample10\Template.xlsx"));
            using (var p = new ExcelPackage())
            {
                p.Load(ms, "");
            }
        }
        [Test]
        public void Issue61()
        {
            DataTable table1 = new DataTable("TestTable");
            table1.Columns.Add("name");
            table1.Columns.Add("id");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("i61");
                ws.Cells["A1"].LoadFromDataTable(table1, true);
                //p.SaveAs(new FileInfo(@"c:\temp\issue61.xlsx"));
            }

        }

        [Test] [Explicit]
        public void Issue58()
        {
            var fileInfo = new FileInfo(@"C:\Temp\issue58.xlsx");
            var package = new ExcelPackage(fileInfo);
            if (package.Workbook.Worksheets.Count > 0)
                package.Workbook.Worksheets.Delete("Test");

            var worksheet = package.Workbook.Worksheets.Add("Test");
            worksheet.Cells[1, 1].Value = "Name";
            worksheet.Cells[1, 2].Value = "Address";
            worksheet.Cells[1, 3].Value = "City";
            worksheet.Cells[2, 1].Value = "Esben Rud";
            worksheet.Cells[2, 2].Value = "Enghavevej";
            worksheet.Cells[2, 3].Value = "Odense";
            var r1 = worksheet.ProtectedRanges.Add("Range1", new ExcelAddress(1, 1, 2, 1));
            Console.WriteLine("Range1: " + r1.Name + " " + r1.Address.ToString());
            var r2 = worksheet.ProtectedRanges.Add("Range2", new ExcelAddress(1, 2, 2, 2));
            Console.WriteLine("Range2: " + r2.Name + " " + r2.Address.ToString());
            Console.WriteLine("*Range1: " + r1.Name + " " + r1.Address.ToString());
            Console.WriteLine("*Range2: " + r2.Name + " " + r2.Address.ToString());
            package.Save();
            package.Dispose();
        }
        [Test]
        public void Issue57()
        {
            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells["A1"].LoadFromArrays(Enumerable.Empty<object[]>());
        }
        [Test] [Explicit]
        public void Issue55()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo(@"C:\Temp\bug\monocell.xlsx"));
            ExcelWorksheet ws = pck.Workbook.Worksheets[1];
            Console.WriteLine(ws.Cells["A1"].Text);
        }
        [Test] [Explicit]
        public void Issue51()
        {
            var filename = new FileInfo(@"c:\temp\bug\bug51.xlsx");
            using (ExcelPackage pck = new ExcelPackage(filename))
            {
                var data = pck.Workbook.Worksheets.Add("data");
                data.Cells["A1"].Value = "Product";
                data.Cells["B1"].Value = "Quantity";
                data.Cells["A2"].Value = "Nails";
                data.Cells["B2"].Value = 37;
                data.Cells["A3"].Value = "Hammer";
                data.Cells["B3"].Value = 5;
                data.Cells["A4"].Value = "Saw";
                data.Cells["B4"].Value = 12;

                var dataRange = data.Cells["A1:B4"];

                var pivot = pck.Workbook.Worksheets.Add("pivot");
                var pivotTable = pivot.PivotTables.Add(pivot.Cells["A1"], dataRange, "a&b");
                var tbl = data.Tables.Add(dataRange, "a&c");
                tbl.Name = "_a&c";
                pivotTable.Name = "a&b";
                pck.Save();
            }
        }
        #region Issue 44
        private static string PIVOT_WS_NAME = "Pivot";
        private static string DATA_WS_NAME = "Data";
        [Test] [Explicit]
        public void Issue44()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("el-GR");
            using (ExcelPackage xlp = new ExcelPackage())
            {
                PrepareDoc(xlp);
                GenPivot(xlp);

                FileStream fs = File.Create(@"c:\temp\bug\pivot44.xlsx");
                xlp.SaveAs(fs);
                fs.Close();
            }

        }
        private void PrepareDoc(ExcelPackage xlp)
        {
            //generate date/value pairs for October 2017
            var series = Enumerable.Range(0, 31);
            var data = from x in series
                       select new { d = new DateTime(2017, 10, x + 1), x = x };
            //put data in table
            ExcelWorksheet ws = xlp.Workbook.Worksheets.Add(DATA_WS_NAME);
            int col = 1;
            ws.Cells[1, col++].Value = "Date";
            ws.Cells[1, col].Value = "Value";
            int row = 2;
            foreach (var line in data)
            {
                col = 1;
                ws.Cells[row, col++].Value = line.d;
                ws.Cells[row, col - 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                ws.Cells[row, col].Value = line.x;
                row++;
            }
        }

        private void GenPivot(ExcelPackage xlp)
        {
            ExcelWorksheet ws = xlp.Workbook.Worksheets.Add(PIVOT_WS_NAME);
            ExcelWorksheet srcws = xlp.Workbook.Worksheets[DATA_WS_NAME];
            ExcelPivotTable piv = ws.PivotTables.Add(ws.Cells[1, 1], srcws.Cells[1, 1, 32, 2], "Pivot1");
            piv.DataFields.Add(piv.Fields["Value"]);
            ExcelPivotTableField dt = piv.RowFields.Add(piv.Fields["Date"]);
            dt.AddDateGrouping(eDateGroupBy.Days | eDateGroupBy.Months);
        }
        #endregion
        [Test]
        public void Issue66()
        {

            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Test!");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Formula = "A1";
                var wb = pck.Workbook;
                wb.Names.Add("Name1", ws.Cells["A1:A2"]);
                ws.Names.Add("Name2", ws.Cells["A1"]);
                pck.Save();
                using (var pck2 = new ExcelPackage(pck.Stream))
                {
                    ws = pck2.Workbook.Worksheets["Test!"];

                }
            }
        }
        [Test] [Explicit]
        public void Issue68()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\bug68.xlsx")))
            {
                var ws = pck.Workbook.Worksheets["Sheet1"];
                pck.Workbook.Worksheets.Delete(ws);
                ws = pck.Workbook.Worksheets.Add("Sheet1");
                pck.SaveAs(new FileInfo(@"c:\temp\bug68-2.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue70()
        {
            var documentPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"c:\temp\workbook with comment.xlsx");
            var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"c:\temp\WorkbookWithCommentOutput.xlsx");
            var fileInfo = new FileInfo(documentPath);
            Assert.That(fileInfo.Exists);
            using (var workbook = new ExcelPackage(fileInfo))
            {
                var ws = workbook.Workbook.Worksheets.First();
                ws.DeleteRow(3); // NRE thrown here
                workbook.SaveAs(new FileInfo(outputPath));
            }
        }

        [Test] [Explicit]
        public void Issue100()
        {
            Stream templateFile = new FileStream(@"c:\temp\bug\epplus_drawing_id_issue.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);
            FileStream outputFile = new FileStream(@"c:\temp\bug\epplus_drawing_id_issue_new.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorkbook wb = package.Workbook;
                ExcelWorksheet sh = wb.Worksheets[1];
                System.Drawing.Image img_ = System.Drawing.Image.FromFile(@"C:\temp\img\background.gif");
                ExcelPicture pic = sh.Drawings.AddPicture("logo", img_);
                pic.SetPosition(1, 1);

                package.SaveAs(outputFile);
            }
        }
        [Test] [Explicit]
        public void Issue99()
        {
            var template = @"c:\temp\bug\iss99\Template.xlsx";
            var result = @"c:\temp\bug\iss99\Result.xlsx";
            using (var inStream = File.Open(template, FileMode.Open))
            {
                using (var outStream = File.Open(result, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    using (ExcelPackage xl = new ExcelPackage(outStream, inStream))
                    {
                        xl.Save();
                    }
                }
            }
        }
        [Test] [Explicit]
        public void Issue94()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\iss94\MergedCellsTemplate.xlsx")))
            {
                var ws = package.Workbook.Worksheets.First();
                var copy = package.Workbook.Worksheets.Add("copy", ws);
                package.Workbook.Worksheets.Delete(ws);
                package.SaveAs(new FileInfo(@"c:\temp\bug\iss94\MergedCellsTemplateSaved.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue107()
        {
            using (ExcelPackage epIN = new ExcelPackage(new FileInfo(@"C:\temp\bug\issue107\in.xlsx")))
            using (ExcelPackage epOUT = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivotbug107.xlsx")))
            {
                foreach (ExcelWorksheet sheet in epIN.Workbook.Worksheets)
                {
                    ExcelWorksheet newSheet = epOUT.Workbook.Worksheets.Add(sheet.Name, sheet);
                }
                epIN.Compatibility.IsWorksheets1Based = true;
                epIN.Workbook.Worksheets.Add(epIN.Workbook.Worksheets[1].Name + "-2", epIN.Workbook.Worksheets[1]);
                epIN.Workbook.Worksheets.Add(epIN.Workbook.Worksheets[2].Name + "-2", epIN.Workbook.Worksheets[2]);
                epOUT.Save();
                epIN.SaveAs(new FileInfo(@"C:\temp\bug\pivotbug107-SameWB.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue127()
        {
            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\PivotTableTestCase.xlsx")))
            {
                Assert.That(p.Workbook.Worksheets.Count, Is.EqualTo(2));
                p.SaveAs(new FileInfo(@"C:\temp\bug\PivotTableTestCaseSaved.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue167()
        {
            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\test-Errorworkbook.xlsx")))
            {
                Assert.That(p.Workbook.Worksheets.Count, Is.EqualTo(1));
                p.SaveAs(new FileInfo(@"C:\temp\bug\test-ErrorworkbookSaved.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue155()
        {
            using (var pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");
                string csv = "A\tB\tC\r\n1\t\t\r\n";
                ws.Cells["A1"].LoadFromText(csv, new ExcelTextFormat { Delimiter = '\t' });
                Assert.That(ws.Cells["B2"].Value == null);
                byte[] data = pck.GetAsByteArray();
                string path = @"C:\temp\test.xlsx";
                File.WriteAllBytes(path, data);
            }
        }
        [Test] [Explicit]
        public void Issue173()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo($@"C:\temp\bug\issue173.xlsx")))
            {
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.First();
                var r = ws.Cells["A4"].Text;
            }
        }
        [Test] [Explicit]
        public void Issue176()
        {
            using (var pck = new ExcelPackage(new FileInfo($@"C:\temp\bug\issue176.xlsx")))
            {
                Assert.That(Math.Round(pck.Workbook.Worksheets[1].Cells["A1"].Style.Fill.BackgroundColor.Tint, 5), Is.EqualTo(-0.04999M));
                pck.SaveAs(new FileInfo($@"C:\temp\bug\issue176-saved.xlsx"));
            }
        }

        [Test] [Explicit]
        public void Issue178()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("TestData");
                var format = new ExcelTextFormat();
                format.TextQualifier = '"';
                var txt = "\"BillingMonth\",\"SequenceNumber\",\"Level7Code\"\r\n";
                txt += "\"022018\",\"1\",\"\"\r\n";

                var range = sheet.Cells["A1"].LoadFromText(txt, format, TableStyles.None, true);
                Assert.That(sheet.Cells["C2"].Value, Is.EqualTo(null));
            }
        }
        [Test] [Explicit]
        public void Issue181()
        {
            using (var pck = new ExcelPackage(new FileInfo($@"C:\temp\bug\issue181.xlsx")))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                var a = pck.Workbook.Properties.Author;
                pck.SaveAs(new FileInfo($@"C:\temp\bug\issue181-saved.xlsx"));
            }
        }
        [Test] [Explicit]
        public void Issue10()
        {
            var fi = new FileInfo($@"C:\temp\bug\issue10.xlsx");
            if (fi.Exists)
            {
                fi.Delete();
            }
            using (var pck = new ExcelPackage(fi))
            {
                var ws = pck.Workbook.Worksheets.Add("Pictures");
                int row = 1;
                foreach (var f in Directory.EnumerateFiles(@"c:\temp\addin_temp\Addin\img\open_icon_library-full\icons\ico\16x16\actions\"))
                {
                    var b = new Bitmap(f);
                    var pic = ws.Drawings.AddPicture($"Image{(row + 1) / 2}", b);
                    pic.SetPosition(row, 0, 0, 0);
                    row += 2;
                }
                pck.Save();
            }
        }
        /// <summary>
        /// Creating a new ExcelPackage with an external stream should not dispose of 
        /// that external stream. That is the responsibility of the caller.
        /// Note: This test would pass with EPPlus 4.1.1. In 4.5.1 the line CloseStream() was added
        /// to the ExcelPackage.Dispose() method. That line is redundant with the line before, 
        /// _stream.Close() except that _stream.Close() is only called if the _stream is NOT
        /// an External Stream (and several other conditions).
        /// Note that CloseStream() doesn't do anything different than _stream.Close().
        /// </summary>
        [Test]
        public void Issue184_Disposing_External_Stream()
        {
            // Arrange
            var stream = new MemoryStream();

            using (var excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Issue 184");
                worksheet.Cells[1, 1].Value = "Hello EPPlus!";
                excelPackage.SaveAs(stream);
                // Act
            } // This dispose should not dispose of stream.

            // Assert
            Assert.That(stream.Length > 0);
        }
        [Test]
        public void Issue204()
        {
            using (var pack = new ExcelPackage())
            {
                //create sheets
                var sheet1 = pack.Workbook.Worksheets.Add("Sheet 1");
                var sheet2 = pack.Workbook.Worksheets.Add("Sheet 2");
                //set some default values
                sheet1.Cells[1, 1].Value = 1;
                sheet2.Cells[1, 1].Value = 2;
                //fill the formula
                var formula = string.Format("'{0}'!R1C1", sheet1.Name);

                var cell = sheet2.Cells[2, 1];
                cell.FormulaR1C1 = formula;
                //Formula should remain the same
                Assert.That(formula.ToUpper(), Is.EqualTo(cell.FormulaR1C1.ToUpper()));
            }
        }
        [Test]
        public void Issue170()
        {
            OpenTemplatePackage("print_titles_170.xlsx");
            _pck.Compatibility.IsWorksheets1Based = false;
            ExcelWorksheet sheet = _pck.Workbook.Worksheets[0];

            sheet.PrinterSettings.RepeatColumns = new ExcelAddress("$A:$C");
            sheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$3");

            SaveWorksheet("print_titles_170-Saved.xlsx");
            _pck.Dispose();
        }
        [Test]
        public void Issue172()
        {
            OpenTemplatePackage("quest.xlsx");
            foreach (var ws in _pck.Workbook.Worksheets)
            {
                Console.WriteLine(ws.Name);
            }

            _pck.Dispose();
        }

        [Test]
        public void Issue219()
        {
            OpenTemplatePackage("issueFile.xlsx");
            foreach (var ws in _pck.Workbook.Worksheets)
            {
                Console.WriteLine(ws.Name);
            }

            _pck.Dispose();
        }
        [Test]
        public void Issue234()
        {
            Assert.Throws<InvalidDataException>(() =>
            {
                using (var s = new MemoryStream())
                {
                    var data = Encoding.UTF8.GetBytes("Bad data").ToArray();
                    s.Write(data, 0, data.Length);
                    var package = new ExcelPackage(s);
                }
            });
        }

        [Test]
        public void Issue220()
        {
            OpenPackage("sheetname_pbl.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Deal's History");
            var a = ws.Cells["A:B"];
            ws.AutoFilterAddress = ws.Cells["A1:C3"];
            _pck.Workbook.Names.Add("Test", ws.Cells["B1:D2"]);
            var name = a.WorkSheet;

            var a2 = new ExcelAddress("'Deal''s History'!a1:a3");
            Assert.That(a2.WorkSheet, Is.EqualTo("Deal's History"));
            _pck.Save();
            _pck.Dispose();

        }
        
        [Test]
        public void Issue233()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                //get some test data
                var cars = Car.GenerateList();

                OpenPackage("issue233.xlsx", true);

                var sheetName = "Summary_GLEDHOWSUGARCO![]()PTY";

                //Create the worksheet 
                var sheet = _pck.Workbook.Worksheets.Add(sheetName);

                //Read the data into a range
                var range = sheet.Cells["A1"].LoadFromCollection(cars, true);

                //Make the range a table
                var tbl = sheet.Tables.Add(range, $"data{sheetName}");
                tbl.ShowTotal = true;
                tbl.Columns["ReleaseYear"].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;

                //save and dispose
                _pck.Save();
                _pck.Dispose();
            });
        }
        public class Car
        {
            public int Id { get; set; }
            public string Make { get; set; }
            public string Model { get; set; }
            public int ReleaseYear { get; set; }

            public Car(int id, string make, string model, int releaseYear)
            {
                Id = Id;
                Make = make;
                Model = model;
                ReleaseYear = releaseYear;
            }

            internal static List<Car> GenerateList()
            {
                return new List<Car>
            {
				//random data
				new Car(1,"Toyota", "Carolla", 1950),
                new Car(2,"Toyota", "Yaris", 2000),
                new Car(3,"Toyota", "Hilux", 1990),
                new Car(4,"Nissan", "Juke", 2010),
                new Car(5,"Nissan", "Trail Blazer", 1995),
                new Car(6,"Nissan", "Micra", 2018),
                new Car(7,"BMW", "M3", 1980),
                new Car(8,"BMW", "X5", 2008),
                new Car(9,"BMW", "M6", 2003),
                new Car(10,"Merc", "S Class", 2001)
            };
            }
        }
        [Test]
        public void Issue236()
        {
            OpenTemplatePackage("Issue236.xlsx");
            _pck.Workbook.Worksheets["Sheet1"].Cells[7, 10].AddComment("test", "Author");
            SaveWorksheet("Issue236-Saved.xlsx");
        }
        [Test]
        public void Issue228()
        {
            OpenTemplatePackage("Font55.xlsx");
            var ws = _pck.Workbook.Worksheets["Sheet1"];
            var d = ws.Drawings.AddShape("Shape1", eShapeStyle.Diamond);
            ws.Cells["A1"].Value = "tasetraser";
            ws.Cells.AutoFitColumns();
            SaveWorksheet("Font55-Saved.xlsx");
        }
        [Test]
        public void Issue241()
        {
            OpenPackage("issue241", true);
            var wks = _pck.Workbook.Worksheets.Add("test");
            wks.DefaultRowHeight = 35;
            _pck.Save();
            _pck.Dispose();
        }
        [Test]
        public void Issue195()
        {
            var pkg = new OfficeOpenXml.ExcelPackage();
            var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
            var defaultStyle = pkg.Workbook.Styles.CreateNamedStyle("Default");
            defaultStyle.Style.Font.Name = "Arial";
            defaultStyle.Style.Font.Size = 18;
            defaultStyle.Style.Font.UnderLine = true;
            var boldStyle = pkg.Workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
            boldStyle.Style.Font.Color.SetColor(Color.Red);

            Assert.That("Arial", Is.EqualTo(defaultStyle.Style.Font.Name));
            Assert.That(18, Is.EqualTo(defaultStyle.Style.Font.Size));

            Assert.That("Arial", Is.EqualTo(boldStyle.Style.Font.Name));
            Assert.That(18, Is.EqualTo(boldStyle.Style.Font.Size));
            Assert.That(boldStyle.Style.Font.Color.Rgb, Is.EqualTo("FFFF0000"));

            pkg.SaveAs(new FileInfo(@"c:\temp\n.xlsx"));
        }
        [Test]
        public void Issue332()
        {
            InitBase();
            var pkg = OpenPackage("Hyperlink.xlsx", true);
            var ws = pkg.Workbook.Worksheets.Add("Hyperlink");
            ws.Cells["A1"].Hyperlink = new ExcelHyperLink("A2", "A2");
            pkg.Save();
        }
        [Test]
        public void Issue332_2()
        {
            InitBase();
            var pkg = OpenPackage("Hyperlink.xlsx");
            var ws = pkg.Workbook.Worksheets["Hyperlink"];
            Assert.That(ws.Cells["A1"].Hyperlink, Is.Not.Null);
        }
        [Test]
        public void Issuer246()
        {
            InitBase();
            var pkg = OpenPackage("issue246.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("DateFormat");
            ws.Cells["A1"].Value = 43465;
            ws.Cells["A1"].Style.Numberformat.Format = @"[$-F800]dddd,\ mmmm\ dd,\ yyyy";
            _pck.Save();

            pkg = OpenPackage("issue246.xlsx");
            ws = _pck.Workbook.Worksheets["DateFormat"];
            var pCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-Se");
            Assert.That(ws.Cells["A1"].Text, Is.EqualTo("den 31 december 2018"));
            Assert.Equals(ws.GetValue<DateTime>(1, 1), new DateTime(2018, 12, 31));
            System.Threading.Thread.CurrentThread.CurrentCulture = pCulture;
        }
        [Test]
        public void Issue347()
        {
            var package = OpenTemplatePackage("Issue327.xlsx");
            var templateWS = package.Workbook.Worksheets["Template"];
            //package.Workbook.Worksheets.Add("NewWs", templateWS);
            package.Workbook.Worksheets.Delete(templateWS);
        }
        [Test]
        public void Issue348()
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("S1");
                string formula = "VLOOKUP(C2,A:B,1,0)";
                ws.Cells[2, 4].Formula = formula;
                var t1 = ws.Cells[2, 4].FormulaR1C1; // VLOOKUP(C2,C[-3]:C[-2],1,0)
                ws.Cells[2, 5].FormulaR1C1 = ws.Cells[2, 4].FormulaR1C1;
                var t2 = ws.Cells[2, 5].FormulaR1C1; // VLOOKUP(C2,C[-3]**:B:C:C**,1,0)   //unexpected value here
            }
        }

        [Test] [Explicit]
        public void Issue376()
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
                var v = ws.DataValidations.AddListValidation("A1");
                v.Formula.Values.Add("1");
                v.Formula.Values.Add("2");
                v.ShowErrorMessage = true;
                v.Error = "error!";
                v.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
                v.AllowBlank = false;
                pck.SaveAs(new FileInfo(@"c:\temp\book.xlsx"));
            }
        }

        [Test]
        public void Issue367()
        {
            using (var pck = OpenTemplatePackage(@"ProductFunctionTest.xlsx"))
            {
                var sheet = pck.Workbook.Worksheets.First();
                //sheet.Cells["B13"].Value = null;
                sheet.Cells["B14"].Value = 11;
                sheet.Cells["B15"].Value = 13;
                sheet.Cells["B16"].Formula = "Product(B13:B15)";
                sheet.Calculate();

                Assert.That(0d, Is.EqualTo(sheet.Cells["B16"].Value));
            }
        }
        [Test]
        public void Issue345()
        {
            using (ExcelPackage package = OpenTemplatePackage("issue345.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets["test"];
                int[] sortColumns = new int[1];
                sortColumns[0] = 0;
                worksheet.Cells["A2:A30864"].Sort(sortColumns);
                package.Save();
            }
        }
        [Test]
        public void Issue387()
        {

            using (ExcelPackage package = OpenTemplatePackage("issue345.xlsx"))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.Add("One");

                worksheet.Cells[1, 3].Value = "Hello";
                var cells = worksheet.Cells["A3"];

                worksheet.Names.Add("R0", cells);
                workbook.Names.Add("Q0", cells);
            }
        }
        [Test]
        public void Issue333()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TextBug");
                ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
                ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

                Assert.That("2019-03-07", Is.EqualTo(ws.Cells["A1"].Text));
            }
        }
        [Test]
        public void Issue445()
        {
            ExcelPackage p = new ExcelPackage();
            ExcelWorksheet ws = p.Workbook.Worksheets.Add("AutoFit"); //<-- This line takes forever. The process hangs.
            ws.Cells[1, 1].Value = new string('a', 50000);
            ws.Cells[1, 1].AutoFitColumns();
        }
        [Test]
        public void Issue460()
        {
            var p = OpenTemplatePackage("Issue460.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var newWs=p.Workbook.Worksheets.Add("NewSheet");
            ws.Cells.Copy(newWs.Cells);
            SaveWorksheet("Issue460_saved.xlsx");
        }
        [Test]
        public void Issue476()
        {
            var p = OpenTemplatePackage("Issue345.xlsx");
            var ws = p.Workbook.Worksheets[0];
            int[] sortColumns = new int[1];
            sortColumns[0] = 0;
            ws.Cells["A2:A64515"].Sort(sortColumns);
            SaveWorksheet("Issue345_saved.xlsx");
        }
    }
}