Attribute VB_Name = "modMainLogic"
Option Explicit
Sub OkEvents()
    Application.EnableEvents = True
End Sub
'Get Info for Drive
Function GetRootDrive(Optional aPath As String) As String
    GetRootDrive = CreateObject("Scripting.FileSystemObject").GetDriveName(aPath)
End Function
Function IsInArray(valueToFind As Variant, arr As Variant) As Boolean
  On Error Resume Next
  IsInArray = Not IsError(Application.Match(valueToFind, arr, 0))
End Function
Function GetUniqueValues(ByRef inputArray() As String) As Collection
    Dim dict As Object
    Dim i As Long
    Set dict = CreateObject("Scripting.Dictionary")

    For i = LBound(inputArray) To UBound(inputArray)
        If Trim(inputArray(i)) <> "" Then
            dict(Trim(inputArray(i))) = 1
        End If
    Next i

    Set GetUniqueValues = New Collection
    Dim key As Variant
    For Each key In dict.Keys
        GetUniqueValues.Add key
    Next key
End Function
Function IFZERO(value As Variant, fallback As Variant) As Variant
    If IsError(value) Then
        IFZERO = CVErr(xlErrValue)
    ElseIf IsNumeric(value) Then
        If CDbl(value) = 0 Then
            IFZERO = fallback
        Else
            IFZERO = value
        End If
    ElseIf Trim(CStr(value)) = "" Then
        IFZERO = fallback
    Else
        IFZERO = value
    End If
End Function

' Utility function to reset user warning preferences
' Call this from the VBA immediate window: ResetUserWarningPreferences
Public Sub ResetUserWarningPreferences()
    UserSettings.ResetHeaderMismatchWarning
    MsgBox "User warning preferences have been reset. All warnings will now be shown again.", vbInformation, "Preferences Reset"
End Sub

' Utility function to show current warning settings
' Call this from the VBA immediate window: ShowUserWarningSettings
Public Sub ShowUserWarningSettings()
    Dim hideHeaderWarning As Boolean
    hideHeaderWarning = UserSettings.GetHideHeaderMismatchWarning()
    
    Dim message As String
    If hideHeaderWarning Then
        message = "Header mismatch warnings are currently HIDDEN." & vbCrLf & vbCrLf & "To show them again, run: ResetUserWarningPreferences"
    Else
        message = "Header mismatch warnings are currently SHOWN (default behavior)."
    End If
    
    MsgBox message, vbInformation, "Current Warning Settings"
End Sub






