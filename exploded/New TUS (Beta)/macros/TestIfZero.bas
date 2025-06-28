Attribute VB_Name = "TestIfZero"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Returns fallback when input is numeric 0")
Private Sub Test_IFZERO_ReturnsFallback_WhenZero()
    Dim result As Variant
    result = IFZERO(0, "fallback")
    Assert.AreEqual "fallback", result
End Sub

'@TestMethod("Returns value when input is non-zero")
Private Sub Test_IFZERO_ReturnsValue_WhenNonZero()
    Dim result As Variant
    result = IFZERO(42, "fallback")
    Assert.AreEqual 42, result
End Sub

'@TestMethod("Returns fallback when coercible text is zero")
Private Sub Test_IFZERO_StringNumericZero()
    Dim result As Variant
    result = IFZERO("0", "fallback")
    Assert.AreEqual "fallback", result
End Sub

'@TestMethod("Returns fallback when input is blank")
Private Sub Test_IFZERO_BlankInput()
    Dim result As Variant
    result = IFZERO("", "fallback")
    Assert.AreEqual "fallback", result ' pass-through
End Sub

'@TestMethod("Returns fallback when coerced input is zero")
Private Sub Test_IFZERO_CoercedFromFormula()
    Dim result As Variant
    result = IFZERO(CDbl(0), "fallback")
    Assert.AreEqual "fallback", result
End Sub

'@TestMethod("Raises #VALUE! error when input is an error")
Private Sub Test_IFZERO_InputsError()
    Dim result As Variant
    result = IFZERO(CVErr(xlErrDiv0), "fallback")
    Assert.AreEqual CVErr(xlErrValue), result
End Sub
