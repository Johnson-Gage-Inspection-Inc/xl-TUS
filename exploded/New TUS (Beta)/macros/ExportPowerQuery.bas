Attribute VB_Name = "ExportPowerQuery"
Option Explicit

Sub ExportAllQueryMCode()
    Dim q As WorkbookQuery
    Dim fso As Object, f As Object, regEx As Object
    Dim exportPath As String, mCode As String, redactedLine As String
    Dim pattern As String, fileBaseName As String
    fileBaseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    exportPath = ThisWorkbook.path & "\exploded\" & fileBaseName & "\queries"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.pattern = "Authorization\s*=\s*""Api-Token [^""]+"""

    For Each q In ThisWorkbook.Queries
        mCode = q.Formula
        mCode = regEx.Replace(mCode, "Authorization = ""Api-Token REDACTED""")

        Set f = fso.CreateTextFile(exportPath & "\" & q.Name & ".m", True, True)
        f.Write mCode
        f.Close
        Debug.Print q.Name & " exported"
    Next q
End Sub


