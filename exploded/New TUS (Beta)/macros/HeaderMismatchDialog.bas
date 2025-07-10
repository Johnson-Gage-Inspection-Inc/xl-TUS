Attribute VB_Name = "HeaderMismatchDialog"
Option Explicit

' Show header mismatch warning with "Don't show again" option
' Returns True if user clicked OK, False if cancelled
Public Function ShowHeaderMismatchWarning(title As String, message As String) As Boolean
    Dim result As VbMsgBoxResult
    
    ' Show the warning with enhanced message
    result = MsgBox(message & vbCrLf & vbCrLf & _
                   "Would you like to hide this type of warning in the future?" & vbCrLf & _
                   "(You can still see mapping info in the immediate window)" & vbCrLf & vbCrLf & _
                   "• Click YES to hide future warnings" & vbCrLf & _
                   "• Click NO to continue showing warnings" & vbCrLf & _
                   "• Click CANCEL to skip this operation", _
                   vbYesNoCancel + vbExclamation, title)
    
    Select Case result
        Case vbYes
            ' Hide future warnings and continue
            UserSettings.SetHideHeaderMismatchWarning True
            ShowHeaderMismatchWarning = True
        Case vbNo
            ' Continue showing warnings
            ShowHeaderMismatchWarning = True
        Case vbCancel
            ' Cancel the operation
            ShowHeaderMismatchWarning = False
    End Select
End Function

