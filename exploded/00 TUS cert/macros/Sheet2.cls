VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("D19, G19, J19, M19, D31, G31, J31, M31")) Is Nothing Then Exit Sub

    On Error GoTo HandleError
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim i1, i2, i3, i4, i5, i6, i7, i8 As Range
    Dim r1, r2, r3, r4, r5, r6, r7, r8 As Range
    
    Set i1 = Range("D19")
    Set i2 = Range("G19")
    Set i3 = Range("J19")
    Set i4 = Range("M19")
    Set i5 = Range("D31")
    Set i6 = Range("G31")
    Set i7 = Range("J31")
    Set i8 = Range("M31")
    
    Set r1 = Range("D10:D18")
    Set r2 = Range("G10:G18")
    Set r3 = Range("J10:J18")
    Set r4 = Range("M10:M18")
    Set r5 = Range("D22:D30")
    Set r6 = Range("G22:G30")
    Set r7 = Range("J22:J30")
    Set r8 = Range("M22:M30")
    
    If i1 = 0 Then
        r1.NumberFormat = "0°F"
    Else
        r1.NumberFormat = "0.0°F"
    End If
    
    If i2 = 0 Then
        r2.NumberFormat = "0°F"
    Else
        r2.NumberFormat = "0.0°F"
    End If

    If i3 = 0 Then
        r3.NumberFormat = "0°F"
    Else
        r3.NumberFormat = "0.0°F"
    End If
    
    If i4 = 0 Then
        r4.NumberFormat = "0°F"
    Else
        r4.NumberFormat = "0.0°F"
    End If
    
    If i5 = 0 Then
        r5.NumberFormat = "0°F"
    Else
        r5.NumberFormat = "0.0°F"
    End If
    
    If i6 = 0 Then
        r6.NumberFormat = "0°F"
    Else
        r6.NumberFormat = "0.0°F"
    End If

    If i7 = 0 Then
        r7.NumberFormat = "0°F"
    Else
        r7.NumberFormat = "0.0°F"
    End If
    
    If i8 = 0 Then
        r8.NumberFormat = "0°F"
    Else
        r8.NumberFormat = "0.0°F"
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

