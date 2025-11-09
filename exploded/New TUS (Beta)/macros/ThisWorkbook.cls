VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Protect sheets with UserInterfaceOnly:=True
    ProtectAllSheets
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Main")
    Dim savePath As Variant

    ProtectAllSheets

    If LCase(Right(ThisWorkbook.Name, 5)) = ".xltm" Then
        With ThisWorkbook.Sheets("Main").Range("A1")
            .value = "Ver. " & Format(Now, "mm.dd.yy")
        End With
        If ThisWorkbook.VBProject.Protection = 0 Then
            ExportVisualBasicCode
        End If
        ExportAllQueryMCode
    
        GoHome
    ElseIf SaveAsUI Then
        Dim suggestedFilename As String: suggestedFilename = ws.Range("G1").value
        If suggestedFilename <> "" Then
            Application.EnableEvents = False ' Prevent recursion
            
            ' Show the Save As dialog with .xlsm format
            savePath = Application.GetSaveAsFilename(InitialFileName:=suggestedFilename, _
                                                     FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
                                                     title:="Save As")

            ' If user selects a file and clicks Save, save in .xlsm format
            If savePath <> False Then ThisWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            
            Application.EnableEvents = True ' Re-enable events
            Cancel = True ' Prevent default Save dialog
        End If
    End If
    ' And, just in case....
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Ensure events are re-enabled and cancel the save operation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Cancel = True
    
    ' Optional: Display error message to user
    MsgBox "An error occurred during the save process: " & Err.Description, vbCritical, "Save Error"
End Sub
Private Sub GoHome()
    Sheets("Main").Select
    Range("D3").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
Private Sub ProtectAllSheets()
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="JGIPyro"
        If IsViewOnlySheet(ws) Then
            'ws.Visible = xlSheetVeryHidden  ' Debug
            ws.Visible = xlSheetHidden
        Else
            ws.Protect Password:="JGIPyro", _
                        UserInterfaceOnly:=True, _
                        AllowFiltering:=True, _
                        AllowSorting:=False, _
                        AllowUsingPivotTables:=False
        End If
    Next ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



