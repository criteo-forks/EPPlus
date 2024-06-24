using System;
using System.Diagnostics.CodeAnalysis;
using NUnit.Framework;
using NUnit.Framework.Constraints;

namespace EPPlusTest;

public static class Assert
{
    public static void AreEquals<T, U>(T left, U right)
    {
        NUnit.Framework.Assert.That(left, Is.EqualTo(right));
    }

    public static void That<TActual>(
        TActual actual,
        IResolveConstraint expression,
        NUnitString message = default(NUnitString))
    {
        NUnit.Framework.Assert.That(actual, expression, message);
    }
    
    public static void That(bool value, NUnitString message = default(NUnitString))
    {
        NUnit.Framework.Assert.That(value, message);
    }
    
    public static TActual? Throws<TActual>(TestDelegate code) where TActual : Exception
    {
        return NUnit.Framework.Assert.Throws<TActual>(code);
    }

    [DoesNotReturn]
    public static void Inconclusive(string message)
    {
        NUnit.Framework.Assert.Inconclusive(message);
    }

    public static void Fail(string message = "")
    {
        NUnit.Framework.Assert.Fail(message);
    }
}