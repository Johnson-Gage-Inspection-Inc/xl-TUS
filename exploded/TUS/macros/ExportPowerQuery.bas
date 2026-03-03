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

    Dim connHeader As String
    Dim conn As WorkbookConnection
    Dim oledb As OLEDBConnection

    For Each q In ThisWorkbook.Queries
        mCode = q.Formula
        mCode = regEx.Replace(mCode, "Authorization = ""Api-Token REDACTED""")

        ' Build a header with connection properties (Usage tab)
        connHeader = ""
        Set conn = Nothing
        Set oledb = Nothing

        On Error Resume Next
        Set conn = ThisWorkbook.Connections("Query - " & q.Name)
        On Error GoTo 0

        If Not conn Is Nothing Then
            Set oledb = conn.OLEDBConnection

            connHeader = "// Connection Properties (Usage tab)" & vbLf
            connHeader = connHeader & "//   BackgroundQuery:       " & oledb.BackgroundQuery & vbLf
            connHeader = connHeader & "//   RefreshOnFileOpen:     " & oledb.RefreshOnFileOpen & vbLf
            connHeader = connHeader & "//   RefreshPeriod:         " & oledb.RefreshPeriod & vbLf
            connHeader = connHeader & "//   RefreshWithRefreshAll: " & conn.RefreshWithRefreshAll & vbLf
            connHeader = connHeader & vbLf
        End If

        Set f = fso.CreateTextFile(exportPath & "\" & q.Name & ".m", True, False)
        f.Write connHeader & mCode
        f.Close
        Debug.Print q.Name & " exported"
    Next q
End Sub
