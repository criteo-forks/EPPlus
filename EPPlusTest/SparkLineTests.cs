using System;
using NUnit.Framework;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;

namespace EPPlusTest
{
    [TestFixture]
    public class SparkLineTests : TestBase
    {
        string _pckfile;
        public SparkLineTests()
        {
            InitBase();
            _pckfile = "Sparklines.xlsx";
        }
        [Test]
        public void StartTest()
        {
            WriteSparklines();
            ReadSparklines();
        }
        public void ReadSparklines()
        {
            _pck = new ExcelPackage();
            OpenPackage(_pckfile);
            var ws = _pck.Workbook.Worksheets[_pck.Compatibility.IsWorksheets1Based?1:0];
            Assert.That(4, Is.EqualTo(ws.SparklineGroups.Count));
            var sg1 = ws.SparklineGroups[0];
            Assert.Equals("A1:A4",sg1.LocationRange.Address);
            Assert.That("B1:C4", Is.EqualTo(sg1.DataRange.Address));
            Assert.That(sg1.DateAxisRange, Is.Null);

            var sg2 = ws.SparklineGroups[1];
            Assert.That("D1:D2", Is.EqualTo(sg2.LocationRange.Address));
            Assert.That("B1:C4", Is.EqualTo(sg2.DataRange.Address));

            var sg3 = ws.SparklineGroups[2];
            Assert.That("A10:B10", Is.EqualTo(sg3.LocationRange.Address));
            Assert.That("B1:C4", Is.EqualTo(sg3.DataRange.Address));

            var sg4 = ws.SparklineGroups[3];
            Assert.That("D10:G10", Is.EqualTo(sg4.LocationRange.Address));
            Assert.That("B1:C4", Is.EqualTo(sg4.DataRange.Address));
            Assert.That("'Sparklines'!A20:A23", Is.EqualTo(sg4.DateAxisRange.Address));

            var c1 = sg1.ColorMarkers;
            Assert.That(c1.Rgb, Is.EqualTo("FFD00000"));
            var ec = sg1.DisplayEmptyCellsAs;
            Assert.That(eDispBlanksAs.Gap, Is.EqualTo(ec));
            var t = sg1.Type;
        }
        public void WriteSparklines()
        {            
            var ws = _pck.Workbook.Worksheets.Add("Sparklines");
            ws.Cells["B1"].Value = 15;
            ws.Cells["B2"].Value = 30;
            ws.Cells["B3"].Value = 35;
            ws.Cells["B4"].Value = 28;
            ws.Cells["C1"].Value = 7;
            ws.Cells["C2"].Value = 33;
            ws.Cells["C3"].Value = 12;
            ws.Cells["C4"].Value = -1;

            //Column<->Row
            var sg1 = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A4"], ws.Cells["B1:C4"]);
            sg1.DisplayEmptyCellsAs = eDispBlanksAs.Gap;
            sg1.Type = eSparklineType.Line;

            //Column<->Column
            var sg2 = ws.SparklineGroups.Add(eSparklineType.Column, ws.Cells["D1:D2"], ws.Cells["B1:C4"]);

            //Row<->Column
            var sg3 = ws.SparklineGroups.Add(eSparklineType.Stacked, ws.Cells["A10:B10"], ws.Cells["B1:C4"]);
            sg3.RightToLeft=true;
            //Row<->Row
            var sg4 = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["D10:G10"], ws.Cells["B1:C4"]);
            ws.Cells["A20"].Value = new DateTime(2016, 12, 30);
            ws.Cells["A21"].Value = new DateTime(2017, 1, 31);
            ws.Cells["A22"].Value = new DateTime(2017, 2, 28);
            ws.Cells["A23"].Value = new DateTime(2017, 3, 31);

            sg4.DateAxisRange = ws.Cells["A20:A23"];

            sg4.ManualMax = 5;
            sg4.ManualMin = 3;

            SaveWorksheet(_pckfile);
        }
    }
}
