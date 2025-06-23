VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub

Public Sub RefreshWireCorrectionData()

    On Error GoTo HandleError

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Refresh Power Query connection (loads corrected offsets)
    ThisWorkbook.Connections("Query - DaqbookOffsets").Refresh

    ' Optional: do post-processing like formatting or export
    ' If Sheet14.Write_Wire_Correction_Factors is still needed, leave this in:
    Call Sheet14.Write_Wire_Correction_Factors

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & vbCrLf & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit

End Sub

