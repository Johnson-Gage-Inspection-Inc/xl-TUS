Attribute VB_Name = "UserSettings"
Option Explicit

' User preference constants
Private Const SETTINGS_SHEET_NAME As String = "UserSettings"
Private Const HIDE_HEADER_MISMATCH_KEY As String = "HideHeaderMismatchWarning"

' Get user setting for hiding header mismatch warnings
Public Function GetHideHeaderMismatchWarning() As Boolean
    On Error GoTo DefaultValue
    
    Dim settingsSheet As Worksheet
    Set settingsSheet = GetOrCreateSettingsSheet()
    
    Dim settingCell As Range
    Set settingCell = FindSettingCell(settingsSheet, HIDE_HEADER_MISMATCH_KEY)
    
    If settingCell Is Nothing Then
        GetHideHeaderMismatchWarning = False
    Else
        GetHideHeaderMismatchWarning = CBool(settingCell.Offset(0, 1).value)
    End If
    Exit Function
    
DefaultValue:
    GetHideHeaderMismatchWarning = False
End Function

' Set user setting for hiding header mismatch warnings
Public Sub SetHideHeaderMismatchWarning(hideWarning As Boolean)
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    Set settingsSheet = GetOrCreateSettingsSheet()
    
    Dim settingCell As Range
    Set settingCell = FindSettingCell(settingsSheet, HIDE_HEADER_MISMATCH_KEY)
    
    If settingCell Is Nothing Then
        ' Add new setting
        Dim lastRow As Long
        lastRow = settingsSheet.Cells(settingsSheet.Rows.count, 1).End(xlUp).Row + 1
        settingsSheet.Cells(lastRow, 1).value = HIDE_HEADER_MISMATCH_KEY
        settingsSheet.Cells(lastRow, 2).value = hideWarning
    Else
        ' Update existing setting
        settingCell.Offset(0, 1).value = hideWarning
    End If
    Exit Sub
    
ErrorHandler:
    ' Silently fail - settings are non-critical
End Sub

' Reset the header mismatch warning setting (show warnings again)
Public Sub ResetHeaderMismatchWarning()
    SetHideHeaderMismatchWarning False
End Sub

' Clear all user settings (for debugging/reset purposes)
Public Sub ClearAllSettings()
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0
    
    If Not settingsSheet Is Nothing Then
        Application.DisplayAlerts = False
        settingsSheet.Delete
        Application.DisplayAlerts = True
    End If
    Exit Sub
    
ErrorHandler:
    ' Silently fail - settings are non-critical
End Sub

' Get or create the hidden settings sheet
Private Function GetOrCreateSettingsSheet() As Worksheet
    Dim ws As Worksheet
    
    ' Try to find existing settings sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new hidden settings sheet
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = SETTINGS_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        
        ' Add headers
        ws.Cells(1, 1).value = "Setting"
        ws.Cells(1, 2).value = "Value"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 2).Font.Bold = True
    End If
    
    Set GetOrCreateSettingsSheet = ws
End Function

' Find a setting cell by key
Private Function FindSettingCell(settingsSheet As Worksheet, settingKey As String) As Range
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = settingsSheet.Cells(settingsSheet.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Start from row 2 (skip header)
        If settingsSheet.Cells(i, 1).value = settingKey Then
            Set FindSettingCell = settingsSheet.Cells(i, 1)
            Exit Function
        End If
    Next i
    
    Set FindSettingCell = Nothing
End Function
