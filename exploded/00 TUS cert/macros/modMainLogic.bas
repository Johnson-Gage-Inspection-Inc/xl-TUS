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

