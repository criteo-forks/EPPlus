using System;
using System.ComponentModel;
using System.Globalization;
using NUnit.Framework;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace EPPlusTest.Utils
{
	[TestFixture]
	public class ConvertUtilTest
	{
		[Test]
		public void TryParseNumericString()
		{
			double result;
			object numericString = null;
			double expected = 0;
			Assert.That(!ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			expected = 1442.0;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // 1.442E+003
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // 1442
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			numericString = expected.ToString("f2", CultureInfo.CurrentCulture); // 1442.00
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			numericString = expected.ToString("n", CultureInfo.CurrentCulture); // 1,442.0
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			expected = -0.00526;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // -5.26E-003
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // -0
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(0.0, Is.EqualTo(result));
			numericString = expected.ToString("f3", CultureInfo.CurrentCulture); // -0.005
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(-0.005, Is.EqualTo(result));
			numericString = expected.ToString("n6", CultureInfo.CurrentCulture); // -0.005260
			Assert.That(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.That(expected, Is.EqualTo(result));
		}
		
		[Test]
		public void TryParseDateString()
		{
			DateTime result;
			object dateString = null;
			DateTime expected = DateTime.MinValue;
			Assert.That(!ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			expected = new DateTime(2013, 1, 15);
			dateString = expected.ToString("d", CultureInfo.CurrentCulture); // 1/15/2013
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			dateString = expected.ToString("D", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 12:00:00 AM
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 12:00 AM
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			expected = new DateTime(2013, 1, 15, 15, 26, 32);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 3:26:32 PM
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(expected, Is.EqualTo(result));
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 3:26 PM
			Assert.That(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.That(new DateTime(2013, 1, 15, 15, 26, 0), Is.EqualTo(result));
		}

        [Test]
        public void TextToInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int>("204");

            Assert.That(204, Is.EqualTo(result));
        }
        // This is just illustration of the bug in old implementation
        //[Test]
        public void TextToIntInOldImplementation()
        {
            var result = GetTypedValue<int>("204");

            Assert.That(204, Is.EqualTo(result));
        }
        [Test]
        public void DoubleToNullableInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int?>(2D);

            Assert.That(2, Is.EqualTo(result));
        }

        [Test]
        public void StringToDecimal()
        {
            var decimalSign=System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var result = ConvertUtil.GetTypedCellValue<decimal>($"1{decimalSign}4");

            Assert.That((decimal)1.4, Is.EqualTo(result));
        }

        [Test]
        public void EmptyStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>("");
            Assert.That(result, Is.Null);
        }

        [Test]
        public void BlankStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>(" ");

            Assert.That(result, Is.Null);
        }

        [Test]
        public void EmptyStringToDecimal()
        {
            Assert.Throws<FormatException>(() =>
            {
                ConvertUtil.GetTypedCellValue<decimal>("");
            });
        }

        [Test]
        public void FloatingPointStringToInt()
        {
            Assert.Throws<FormatException>(() =>
            {
                ConvertUtil.GetTypedCellValue<int>("1.4");
            });
        }

        [Test]
        public void IntToDateTime()
        {
            Assert.Throws<InvalidCastException>(() =>
            {
                ConvertUtil.GetTypedCellValue<DateTime>(122);
            });
            
        }

        [Test]
        public void IntToTimeSpan()
        {
            Assert.Throws<InvalidCastException>(() =>
            {
                ConvertUtil.GetTypedCellValue<TimeSpan>(122);
            });
        }

        [Test]
        public void IntStringToTimeSpan()
        {
            Assert.That(TimeSpan.FromDays(122), Is.EqualTo(ConvertUtil.GetTypedCellValue<TimeSpan>("122")));
        }

        [Test]
        public void BoolToInt()
        {
            Assert.That(1, Is.EqualTo(ConvertUtil.GetTypedCellValue<int>(true)));
            Assert.That(0, Is.EqualTo(ConvertUtil.GetTypedCellValue<int>(false)));
        }

        [Test]
        public void BoolToDecimal()
        {
            Assert.That(1m, Is.EqualTo(ConvertUtil.GetTypedCellValue<decimal>(true)));
            Assert.That(0m, Is.EqualTo(ConvertUtil.GetTypedCellValue<decimal>(false)));
        }

        [Test]
        public void BoolToDouble()
        {
            Assert.That(1d, Is.EqualTo(ConvertUtil.GetTypedCellValue<double>(true)));
            Assert.That(0d, Is.EqualTo(ConvertUtil.GetTypedCellValue<double>(false)));
        }

        [Test]
        public void BadTextToInt()
        {
            Assert.Throws<FormatException>(() =>
            {
                ConvertUtil.GetTypedCellValue<int>("text1");
            });
        }

        // previous implementation
        internal T GetTypedValue<T>(object v)
        {
            if (v == null)
            {
                return default(T);
            }
            Type fromType = v.GetType();
            Type toType = typeof(T);
            
            Type toType2 = (TypeCompat.IsGenericType(toType) && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                ? Nullable.GetUnderlyingType(toType)
                : null;
            if (fromType == toType || fromType == toType2)
            {
                return (T)v;
            }
            var cnv = TypeDescriptor.GetConverter(fromType);
            if (toType == typeof(DateTime) || toType2 == typeof(DateTime))    //Handle dates
            {
                if (fromType == typeof(TimeSpan))
                {
                    return ((T)(object)(new DateTime(((TimeSpan)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    DateTime dt;
                    if (DateTime.TryParse(v.ToString(), out dt))
                    {
                        return (T)(object)(dt);
                    }
                    else
                    {
                        return default(T);
                    }

                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {
                        return (T)(object)(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))));
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
            else if (toType == typeof(TimeSpan) || toType2 == typeof(TimeSpan))    //Handle timespan
            {
                if (fromType == typeof(DateTime))
                {
                    return ((T)(object)(new TimeSpan(((DateTime)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    TimeSpan ts;
                    if (TimeSpan.TryParse(v.ToString(), out ts))
                    {
                        return (T)(object)(ts);
                    }
                    else
                    {
                        return default(T);
                    }
                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {

                        return (T)(object)(new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks));
                    }
                    else
                    {
                        try
                        {
                            // Issue 14682 -- "GetValue<decimal>() won't convert strings"
                            // As suggested, after all special cases, all .NET to do it's 
                            // preferred conversion rather than simply returning the default
                            return (T)Convert.ChangeType(v, typeof(T));
                        }
                        catch (Exception)
                        {
                            // This was the previous behaviour -- no conversion is available.
                            return default(T);
                        }
                    }
                }
            }
            else
            {
                if (cnv.CanConvertTo(toType))
                {
                    return (T)cnv.ConvertTo(v, typeof(T));
                }
                else
                {
                    if (toType2 != null)
                    {
                        toType = toType2;
                        if (cnv.CanConvertTo(toType))
                        {
                            return (T)cnv.ConvertTo(v, toType); //Fixes issue 15377
                        }
                    }

                    if (fromType == typeof(double) && toType == typeof(decimal))
                    {
                        return (T)(object)Convert.ToDecimal(v);
                    }
                    else if (fromType == typeof(decimal) && toType == typeof(double))
                    {
                        return (T)(object)Convert.ToDouble(v);
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
        }

    }
}
