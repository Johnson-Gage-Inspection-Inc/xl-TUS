Attribute VB_Name = "SheetGuards"
Option Explicit
Public Sub EnforceSheetIsViewOnly(sh As Worksheet)
    ' Exit Sub ' Remove for production/beta testing
    Dim msg As String
    On Error Resume Next
    msg = "This sheet is protected. Use the interface on the Main sheet."

    If Not Application.UserName Like "*Admin*" Then
        Application.EnableEvents = False
        sh.Range("A1").Select
        MsgBox msg, vbInformation
        Application.EnableEvents = True
    End If
End Sub
Public Function IsViewOnlySheet(ws As Worksheet) As Boolean
    Dim lo As ListObject

    On Error Resume Next

    For Each lo In ws.ListObjects
        If lo.SourceType <> xlSrcRange Then
            IsViewOnlySheet = True
            Exit Function
        End If
    Next lo

    IsViewOnlySheet = False
End Function
