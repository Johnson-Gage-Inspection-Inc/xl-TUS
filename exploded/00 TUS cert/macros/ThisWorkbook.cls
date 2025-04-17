VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Protect sheets with UserInterfaceOnly:=True
    ProtectAllSheets
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ProtectAllSheets

    If LCase(Right(ThisWorkbook.Name, 5)) = ".xltm" Then
        With ThisWorkbook.Sheets("Main").Range("A1")
            .Value = "Ver. " & Format(Now, "mm.dd.yy")
        End With
        If ThisWorkbook.VBProject.Protection = 0 Then
            ExportVisualBasicCode
        End If
        ExportAllQueryMCode
    
        GoHome
    End If

    ' And, just in case....
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Private Sub GoHome()
        Sheets("Main").Select
        Range("D3").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
End Sub
Private Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="JGIPyro"
        If IsViewOnlySheet(ws) Then
            ws.Visible = xlSheetVeryHidden
        Else
            ws.Protect Password:="JGIPyro", _
                        UserInterfaceOnly:=True, _
                        AllowFiltering:=True, _
                        AllowSorting:=False, _
                        AllowUsingPivotTables:=False
        End If
    Next ws
End Sub
