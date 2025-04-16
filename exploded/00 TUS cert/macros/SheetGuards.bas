Attribute VB_Name = "SheetGuards"
Option Explicit
Public Sub EnforceSheetIsViewOnly(sh As Worksheet)
    Dim msg As String
    On Error Resume Next
    msg = "This sheet is protected. Use the interface on the Main sheet."

    ' Basic override logic — adjust to your environment
    If Not Application.UserName Like "*Admin*" Then
        Application.EnableEvents = False
        sh.Range("A1").Select
        MsgBox msg, vbInformation
        Application.EnableEvents = True
    End If
End Sub
Public Function IsViewOnlySheet(ws As Worksheet) As Boolean
    Dim vbCode As String
    Dim comp As Object

    On Error GoTo Fail

    Set comp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    vbCode = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)

    IsViewOnlySheet = (InStr(1, vbCode, "EnforceSheetIsViewOnly", vbTextCompare) > 0)
    Exit Function

Fail:
    IsViewOnlySheet = False
End Function
