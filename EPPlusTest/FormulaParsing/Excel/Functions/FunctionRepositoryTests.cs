using System;
using System.Collections.Generic;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestFixture]
    public class FunctionRepositoryTests
    {
        #region LoadModule Tests
        [Test]
        public void LoadModulePopulatesFunctionsAndCustomCompilers()
        {
            var functionRepository = FunctionRepository.Create();
            Assert.That(!functionRepository.IsFunctionName(MyFunction.Name));
            Assert.That(!functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
            functionRepository.LoadModule(new TestFunctionModule());
            Assert.That(functionRepository.IsFunctionName(MyFunction.Name));
            Assert.That(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
            // Make sure reloading the module overwrites previous functions and compilers
            functionRepository.LoadModule(new TestFunctionModule());
        }
        #endregion

        #region Nested Classes
        public class TestFunctionModule : FunctionsModule
        {
            public TestFunctionModule()
            {
                var myFunction = new MyFunction();
                var customCompiler = new MyFunctionCompiler(myFunction, ParsingContext.Create());
                base.Functions.Add(MyFunction.Name, myFunction);
                base.CustomCompilers.Add(typeof(MyFunction), customCompiler);
            }
        }

        public class MyFunction : ExcelFunction
        {
            public const string Name = "MyFunction";
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }

        public class MyFunctionCompiler : FunctionCompiler
        {
            public MyFunctionCompiler(MyFunction function, ParsingContext context) : base(function, context) { }
            public override CompileResult Compile(IEnumerable<Expression> children)
            {
                throw new NotImplementedException();
            }
        }
        #endregion
    }
}
