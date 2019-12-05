using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace RubberduckTests.Refactoring.EncapsulateField
{

    [TestFixture]
    public class EncapsulateFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("fizz", true, "baz", true, "buzz", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true)]
        [TestCase("fizz", false, "baz", false, "buzz", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields(
            string var1, bool var1Flag,
            string var2, bool var2Flag,
            string var3, bool var3Flag)
        {
            string inputCode =
$@"Public {var1} As Integer
Public {var2} As Integer
Public {var3} As Integer";

            var selection = new Selection(1, 1);

            var userInput = new UserInputDataObject(var1, $"{var1}Prop", var1Flag);
            userInput.AddAttributeSet(var2, $"{var2}Prop", var2Flag);
            userInput.AddAttributeSet(var3, $"{var3}Prop", var3Flag);

            var flags = new Dictionary<string, bool>()
            {
                [var1] = var1Flag,
                [var2] = var2Flag,
                [var3] = var3Flag
            };

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

            var notEncapsulated = flags.Keys.Where(k => !flags[k])
                   .Select(k => k);

            var encapsulated = flags.Keys.Where(k => flags[k])
                   .Select(k => k);

            foreach ( var variable in notEncapsulated)
            {
                StringAssert.Contains($"Public {variable} As Integer", actualCode);
            }

            foreach (var variable in encapsulated)
            {
                StringAssert.Contains($"Private {variable} As", actualCode);
                StringAssert.Contains($"{variable}Prop = {variable}", actualCode);
                StringAssert.Contains($"{variable} = value", actualCode);
                StringAssert.Contains($"Let {variable}Prop(ByVal value As", actualCode);
                StringAssert.Contains($"Property Get {variable}Prop()", actualCode);
            }
        }

        [TestCase("fizz", true, "baz", true, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true, "boink", false)]
        [TestCase("fizz", false, "baz", true, "buzz", false, "boink", false)]
        [TestCase("fizz", false, "baz", false, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", false, "buzz", false, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", false, "boink", true)]
        [TestCase("fizz", false, "baz", false, "buzz", true, "boink", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFieldsInList(
            string var1, bool var1Flag,
            string var2, bool var2Flag,
            string var3, bool var3Flag,
            string var4, bool var4Flag)
        {
            string inputCode =
$@"Public {var1} As Integer, {var2} As Integer, {var3} As Integer, {var4} As Integer";

            var selection = new Selection(1, 9);

            var userInput = new UserInputDataObject(var1, $"{var1}Prop", var1Flag);
            userInput.AddAttributeSet(var2, $"{var2}Prop", var2Flag);
            userInput.AddAttributeSet(var3, $"{var3}Prop", var3Flag);
            userInput.AddAttributeSet(var4, $"{var4}Prop", var4Flag);

            var flags = new Dictionary<string, bool>()
            {
                [var1] = var1Flag,
                [var2] = var2Flag,
                [var3] = var3Flag,
                [var4] = var4Flag
            };

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

            var remainInList = flags.Keys.Where(k => !flags[k])
                   .Select(k => $"{k} As Integer");

            if (remainInList.Any())
            {
                var declarationList = $"Public {string.Join(", ", remainInList)}";
                StringAssert.Contains(declarationList, actualCode);
            }
            else
            {
                StringAssert.DoesNotContain($"Public {Environment.NewLine}", actualCode);
            }

            foreach (var key in flags.Keys)
            {
                if (flags[key])
                {
                    StringAssert.Contains($"Private {key} As", actualCode);
                    StringAssert.Contains($"{key}Prop = {key}", actualCode);
                    StringAssert.Contains($"{key} = value", actualCode);
                    StringAssert.Contains($"Let {key}Prop(ByVal value As", actualCode);
                    StringAssert.Contains($"Property Get {key}Prop()", actualCode);
                }
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidDeclarationType_Throws()
        {
            const string inputCode =
                @"Public fizz As Integer";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, presenterAction, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidIdentifierSelected_Throws()
        {
            const string inputCode =
                @"Public Function fiz|z() As Integer
End Function";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var codeString = inputCode.ToCodeString();
            var actualCode = RefactoredCode(codeString.Code, codeString.CaretPosition.ToOneBased(), presenterAction, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(codeString.Code, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldIsOverMultipleLines()
        {
            const string inputCode =
                @"Public _
fi|zz _
As _
Integer";
            const string expectedCode =
                @"Private _
fizz _
As _
Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_ReadOnly()
        {
            const string inputCode =
                @"|Public fizz As Integer";

            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", isReadonly: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_NewPropertiesInsertedAboveExistingCode()
        {
            const string inputCode =
                @"|Public fizz As Integer

Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.Greater(actualCode.IndexOf("Sub Foo"), actualCode.LastIndexOf("End Property"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherPropertiesInClass()
        {
            const string inputCode =
                @"|Public fizz As Integer

Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.Greater(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("fizz As Integer"));
            Assert.Less(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("Get Foo"));
        }

        [TestCase("|Public fizz As Integer\r\nPublic buzz As Boolean", "Private fizz As Integer\r\nPublic buzz As Boolean")]
        [TestCase("Public buzz As Boolean\r\n|Public fizz As Integer", "Public buzz As Boolean\r\nPrivate fizz As Integer")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherNonSelectedFieldsInClass(string inputFields, string expectedFields)
        {
            string inputCode = inputFields;

            string expectedCode =
$@"{expectedFields}

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase(1, 10, "fizz", "Public buzz", "Private fizz As Variant", "Public fizz")]
        [TestCase(2, 2, "buzz", "Public fizz, _\r\nbazz", "Private buzz As Boolean", "")]
        [TestCase(3, 2, "bazz", "Public fizz, _\r\nbuzz", "Private bazz As Date", "Boolean, bazz As Date")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_SelectedWithinDeclarationList(int rowSelection, int columnSelection, string fieldName, string contains1, string contains2, string doesNotContain)
        {
            string inputCode =
$@"Public fizz, _
buzz As Boolean, _
bazz As Date";

            var selection = new Selection(rowSelection, columnSelection);
            var presenterAction = Support.SetParametersForSingleTarget(fieldName, "Name");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains(contains1, actualCode);
            StringAssert.Contains(contains1, actualCode);
            if (doesNotContain.Length > 0)
            {
                StringAssert.DoesNotContain(doesNotContain, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_NameConflict()
        {
            const string inputCode =
                @"Private fizz As String
Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property
";
            var fieldName = "fizz";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                IEncapsulateFieldCandidate efd = null;
                var fields = new List<IEncapsulateFieldCandidate>();
                var validator = new EncapsulateFieldNamesValidator(state, () => fields);

                var match = state.DeclarationFinder.MatchName(fieldName).Single();
                efd = new EncapsulateFieldCandidate(match, validator);
                fields.Add(efd);
                efd.PropertyName = "Name";

                var hasConflict = !validator.HasValidEncapsulationAttributes(efd.EncapsulationAttributes, efd.QualifiedModuleName, new Declaration[] { efd.Declaration }); //(Declaration dec) => match.Equals(dec));
                Assert.IsTrue(hasConflict);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateFieldAsUDT()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Name As Integer", actualCode);
            StringAssert.Contains("this.Name = value", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_Defaults()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            const string expectedCode =
                @"Private fizz1 As Integer

Public Property Get Fizz() As Integer
    Fizz = fizz1
End Property

Public Property Let Fizz(ByVal value As Integer)
    fizz1 = value
End Property
";
            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_DefaultsAsUDT()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.UserAcceptsDefaults(asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Fizz As Integer", actualCode);
            StringAssert.Contains("this As This_Type", actualCode);
            StringAssert.Contains("this.Fizz = value", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldHasReferences()
        {
            const string inputCode =
                @"|Public fizz As Integer

Sub Foo()
    fizz = 0
    Bar fizz
End Sub

Sub Bar(ByVal name As Integer)
End Sub";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var enapsulationIdentifiers = new EncapsulationIdentifiers("fizz") { Property = "Name" };

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.AreEqualIgnoringCase(enapsulationIdentifiers.Field, "fizz");
            StringAssert.Contains($"Private {enapsulationIdentifiers.Field} As Integer", actualCode);
            StringAssert.Contains("Property Get Name", actualCode);
            StringAssert.Contains("Property Let Name", actualCode);
            StringAssert.Contains($"Name = {enapsulationIdentifiers.Field}", actualCode);
            StringAssert.Contains($"{enapsulationIdentifiers.Field} = value", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void GivenReferencedPublicField_UpdatesReferenceToNewProperty()
        {
            const string codeClass1 =
                @"|Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            const string codeClass2 =
                @"Sub Foo()
    Dim c As Class1
    c.fizz = 0
    Bar c.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", true);

            var class1CodeString = codeClass1.ToCodeString();
            var actualCode = RefactoredCode(
                "Class1",
                class1CodeString.CaretPosition.ToOneBased(), 
                presenterAction, 
                null, 
                false, 
                ("Class1", class1CodeString.Code, ComponentType.ClassModule),
                ("Class2", codeClass2, ComponentType.ClassModule));

            StringAssert.Contains("Name = 1", actualCode["Class1"]);
            StringAssert.Contains("c.Name = 0", actualCode["Class2"]);
            StringAssert.Contains("Bar c.Name", actualCode["Class2"]);
            StringAssert.DoesNotContain("fizz", actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_PresenterIsNull()
        {
            const string inputCode =
                @"Private fizz As Variant";
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), Selection.Home);
                var selectionService = MockedSelectionService();
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(It.IsAny<EncapsulateFieldModel>()))
                    .Returns(() => null); // resolves ambiguous method overload

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_ModelIsNull()
        {
            const string inputCode =
                @"|Private fizz As Variant";

            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction = model => null;

            var codeString = inputCode.ToCodeString();
            var actualCode = Support.RefactoredCode(codeString, presenterAction, typeof(InvalidRefactoringModelException));
            Assert.AreEqual(codeString.Code, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OptionExplicit_NotMoved()
        {
            const string inputCode =
                @"Option Explicit

Public foo As String";

            const string expectedCode =
                @"Option Explicit

Private foo As String

Public Property Get Name() As String
    Name = foo
End Property

Public Property Let Name(ByVal value As String)
    foo = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("foo", "Name");
            var actualCode = RefactoredCode(inputCode, "foo", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void Refactoring_Puts_Code_In_Correct_Place()
        {
            const string inputCode =
                @"Option Explicit

Public Fo|o As String";

            const string expectedCode =
                @"Option Explicit

Private Foo As String

Public Property Get bar() As String
    bar = Foo
End Property

Public Property Let bar(ByVal value As String)
    Foo = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("Foo", "bar");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("Private", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Public", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Private", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Public", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Private", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Public", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Private", "mArray() As String", "mArray() As String")]
        [TestCase("Public", "mArray() As String", "mArray() As String")]
        [TestCase("Private", "mArray(5)", "mArray(5) As Variant")]
        [TestCase("Public", "mArray(5)", "mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray(string visibility, string arrayDeclaration, string expectedArrayDeclaration)
        {
            string inputCode =
                $@"Option Explicit

{visibility} {arrayDeclaration}";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Private {expectedArrayDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var userInput = new UserInputDataObject("mArray", "MyArray", true);

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("5")]
        [TestCase("5,2,3")]
        [TestCase("1 to 100")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_DeclaredInList(string dimensions)
        {
            string inputCode =
                $@"Option Explicit

Public mArray({dimensions}) As String, anotherVar As Long, andOneMore As Variant";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Public anotherVar As Long, andOneMore As Variant
Private mArray({dimensions}) As String

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Public anotherVar As Long, andOneMore As Variant", actualCode);
            StringAssert.Contains($"Private mArray({dimensions}) As String", actualCode);
            StringAssert.Contains("Get MyArray() As Variant", actualCode);
            StringAssert.Contains("MyArray = mArray", actualCode);
            StringAssert.DoesNotContain("Let MyArray", actualCode);
            StringAssert.DoesNotContain("Set MyArray", actualCode);
        }

        [TestCase("mArr|ay(5) As String, mNextVar As Long", "Private mArray(5) As String")]
        [TestCase("mNextVar As Long, mArr|ay(5) As String", "Private mArray(5) As String")]
        [TestCase("mArr|ay(5), mNextVar As Long", "Private mArray(5) As Variant")]
        [TestCase("mNextVar As Long, mAr|ray(5)", "Private mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldNameForFieldInList(string declarationList, string expectedDeclaration)
        {
            string inputCode =
                $@"Option Explicit

Public {declarationList}";

            string expectedCode =
                $@"Option Explicit

Public mNextVar As Long
{expectedDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase(false)]
        [TestCase(true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void StandardModuleSource_ExternalReferences(bool moduleResolve)
        {
            var sourceModuleName = "SourceModule";
            var referenceExpression = moduleResolve ? $"{sourceModuleName}." : string.Empty;
            var sourceModuleCode =
$@"

Public th|is As Long";

            var procedureModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            string classModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            var userInput = new UserInputDataObject("this", "MyProperty", true);

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingModuleCode);

            var referencingClassCode = actualModuleCode["ClassModule"];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingClassCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingClassCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingClassCode);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
//<<<<<<< HEAD
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
//=======
//            var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
//            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
//            var uiDispatcherMock = new Mock<IUiDispatcher>();
//            uiDispatcherMock
//                .Setup(m => m.Invoke(It.IsAny<Action>()))
//                .Callback((Action action) => action.Invoke());
//            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider, uiDispatcherMock.Object);
//>>>>>>> rubberduck-vba/next
        }
    }
}
