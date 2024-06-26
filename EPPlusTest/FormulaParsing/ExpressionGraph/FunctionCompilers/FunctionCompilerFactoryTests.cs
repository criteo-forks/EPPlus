﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace EPPlusTest.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    [TestFixture]
    public class FunctionCompilerFactoryTests
    {
        private ParsingContext _context;

        [SetUp]
        public void Initialize()
        {
            _context = ParsingContext.Create();
        }
        #region Create Tests
        [Test]
        public void CreateHandlesStandardFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new Sum();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<DefaultCompiler>());
        }

        [Test]
        public void CreateHandlesSpecialIfCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new If();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<IfFunctionCompiler>());
        }

        [Test]
        public void CreateHandlesSpecialIfErrorCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IfError();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<IfErrorFunctionCompiler>());
        }

        [Test]
        public void CreateHandlesSpecialIfNaCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IfNa();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<IfNaFunctionCompiler>());
        }

        [Test]
        public void CreateHandlesLookupFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new Column();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<LookupFunctionCompiler>());
        }

        [Test]
        public void CreateHandlesErrorFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IsError();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<ErrorHandlingFunctionCompiler>());
        }

        [Test]
        public void CreateHandlesCustomFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            functionRepository.LoadModule(new TestFunctionModule(_context));
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new MyFunction();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.That(functionCompiler, Is.InstanceOf<MyFunctionCompiler>());
        }
        #endregion

        #region Nested Classes
        public class TestFunctionModule : FunctionsModule
        {
            public TestFunctionModule(ParsingContext context)
            {
                var myFunction = new MyFunction();
                var customCompiler = new MyFunctionCompiler(myFunction, context);
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
