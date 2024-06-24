using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestFixture]
    public class ExcelChartAxisTest
    {
        private ExcelChartAxis axis;
        
        [SetUp]
        public void Initialize()
        {
            var xmlDoc = new XmlDocument();
            var xmlNsm = new XmlNamespaceManager(new NameTable());
            xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
            axis = new ExcelChartAxis(xmlNsm, xmlDoc.CreateElement("axis"));
        }

        [Test]
        public void CrossesAt_SetTo2_Is2()
        {
            axis.CrossesAt = 2;
            Assert.That(axis.CrossesAt, Is.EqualTo(2));
        }

        [Test]
        public void CrossesAt_SetTo1EMinus6_Is1EMinus6()
        {
            axis.CrossesAt = 1.2e-6;
            Assert.That(axis.CrossesAt, Is.EqualTo(1.2e-6));
        }

        [Test]
        public void MinValue_SetTo2_Is2()
        {
            axis.MinValue = 2;
            Assert.That(axis.MinValue, Is.EqualTo(2));
        }

        [Test]
        public void MinValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MinValue = 1.2e-6;
            Assert.That(axis.MinValue, Is.EqualTo(1.2e-6));
        }

        [Test]
        public void MaxValue_SetTo2_Is2()
        {
            axis.MaxValue = 2;
            Assert.That(axis.MaxValue, Is.EqualTo(2));
        }

        [Test]
        public void MaxValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MaxValue = 1.2e-6;
            Assert.That(axis.MaxValue, Is.EqualTo(1.2e-6));
        }
        [Test] 
        public void Gridlines_Set_IsNotNull()
        { 
            var major = axis.MajorGridlines; 
            Assert.That(axis.ExistNode("c:majorGridlines")); 
  
            var minor = axis.MinorGridlines; 
            Assert.That(axis.ExistNode("c:minorGridlines")); 
        } 
  
        [Test] 
        public void Gridlines_Remove_IsNull()
        { 
            var major = axis.MajorGridlines; 
            var minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(); 
  
            Assert.That(!axis.ExistNode("c:majorGridlines")); 
            Assert.That(!axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines; 
            minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(true, false); 
  
            Assert.That(!axis.ExistNode("c:majorGridlines")); 
            Assert.That(axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines; 
            minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(false, true); 
  
            Assert.That(axis.ExistNode("c:majorGridlines")); 
            Assert.That(!axis.ExistNode("c:minorGridlines")); 
        } 
    }
}
