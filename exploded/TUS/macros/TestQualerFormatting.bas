Attribute VB_Name = "TestQualerFormatting"
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

'@TestMethod
Public Sub GetFormattedWorkItemNumber_OptionalItemInput_OmittedOrBlank()

    Dim cases As Variant
    cases = Array( _
        Array("79123", "56561-079123"), _
        Array("079123", "56561-079123"), _
        Array("56561-079123", "56561-079123"), _
        Array("79123.01", "56561-079123.01"), _
        Array("56561-079123.01", "56561-079123.01") _
    )

    Dim i As Long
    For i = LBound(cases) To UBound(cases)
        Dim rawOrder As String: rawOrder = cases(i)(0)
        Dim expected As String: expected = cases(i)(1)

        Dim actualOmitted As String
        actualOmitted = GetFormattedWorkItemNumber(rawOrder)
        If Trim(actualOmitted) <> Trim(expected) Then
            Assert.Fail "FAILED (omitted item): OrderInput='" & rawOrder & "'" & vbCrLf & _
                        "Expected: '" & expected & "'" & vbCrLf & _
                        "Actual:   '" & actualOmitted & "'"
        End If

        Dim actualBlank As String
        actualBlank = GetFormattedWorkItemNumber(rawOrder, "")
        If Trim(actualBlank) <> Trim(expected) Then
            Assert.Fail "FAILED (blank item): OrderInput='" & rawOrder & "'" & vbCrLf & _
                        "Expected: '" & expected & "'" & vbCrLf & _
                        "Actual:   '" & actualBlank & "'"
        End If
    Next i

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

'@TestMethod
Public Sub GetFormattedWorkItemNumber_ExplicitExpectations()

    Dim OrderInputs As Variant
    OrderInputs = Array( _
        Array("12345", "56561-012345"), _
        Array("012345", "56561-012345"), _
        Array("12345.01", "56561-012345.01"), _
        Array("012345.01", "56561-012345.01"), _
        Array("56561-012345", "56561-012345"), _
        Array("56561-012345.01", "56561-012345.01") _
    )

    Dim ItemInputs As Variant
    ItemInputs = Array( _
        Array("1", "01"), _
        Array("01", "01"), _
        Array("01R1", "01R1"), _
        Array("99R99", "99R99") _
    )

    Dim o As Long, i As Long
    For o = LBound(OrderInputs) To UBound(OrderInputs)
        For i = LBound(ItemInputs) To UBound(ItemInputs)
        
            Dim rawOrder As String: rawOrder = OrderInputs(o)(0)
            Dim expectedOrder As String: expectedOrder = OrderInputs(o)(1)

            Dim rawItem As String: rawItem = ItemInputs(i)(0)
            Dim expectedItem As String: expectedItem = ItemInputs(i)(1)

            Dim expected As String: expected = expectedOrder & "-" & expectedItem
            Dim actual As String
            actual = GetFormattedWorkItemNumber(rawOrder, rawItem)

            If Trim(actual) <> Trim(expected) Then
                Assert.Fail "FAILED for: OrderInput='" & rawOrder & "', ItemInput='" & rawItem & "'" & vbCrLf & _
                            "Expected: '" & expected & "'" & vbCrLf & _
                            "Actual:   '" & actual & "'"
            End If

        Next i
    Next o

End Sub


