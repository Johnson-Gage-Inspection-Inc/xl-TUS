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
    Dim fastLoad As String

    For Each q In ThisWorkbook.Queries
        mCode = q.Formula
        mCode = regEx.Replace(mCode, "Authorization = ""Api-Token REDACTED""")

        ' Build a header with connection properties (Usage tab)
        connHeader = ""
        Set conn = Nothing
        Set oledb = Nothing
        fastLoad = ""

        On Error Resume Next
        Set conn = ThisWorkbook.Connections("Query - " & q.Name)
        On Error GoTo 0

        If Not conn Is Nothing Then
            Set oledb = conn.OLEDBConnection

            ' Enable Fast Data Load = NOT PreserveFormatting on the linked QueryTable.
            Dim ws As Worksheet
            Dim lo As ListObject
            Dim result As Variant

            For Each ws In ThisWorkbook.Worksheets
                For Each lo In ws.ListObjects
                    result = GetFastDataLoad(lo, conn)
                    If Not IsEmpty(result) Then
                        fastLoad = CStr(result)
                        GoTo FoundQT
                    End If
                Next lo
            Next ws
FoundQT:
            On Error GoTo 0

            connHeader = "// Connection Properties (Usage tab)" & vbLf
            connHeader = connHeader & "//   BackgroundQuery:       " & oledb.BackgroundQuery & vbLf
            connHeader = connHeader & "//   RefreshOnFileOpen:     " & oledb.RefreshOnFileOpen & vbLf
            connHeader = connHeader & "//   RefreshPeriod:         " & oledb.RefreshPeriod & vbLf
            connHeader = connHeader & "//   RefreshWithRefreshAll: " & conn.RefreshWithRefreshAll & vbLf
            If fastLoad <> "" Then
                connHeader = connHeader & "//   EnableFastDataLoad:    " & fastLoad & vbLf
            End If
            connHeader = connHeader & vbLf
        End If

        Set f = fso.CreateTextFile(exportPath & "\" & q.Name & ".m", True, False)
        f.Write connHeader & mCode
        f.Close
        Debug.Print q.Name & " exported"
    Next q
End Sub

' Returns True/False for EnableFastDataLoad if lo's QueryTable matches conn.
' Returns Empty if the ListObject has no QueryTable or doesn't match.
Private Function GetFastDataLoad(lo As ListObject, conn As WorkbookConnection) As Variant
    GetFastDataLoad = Empty

    ' SourceType 1 = xlSrcRange (plain table, no QueryTable) - skip these
    ' to avoid the non-trappable 1004 from accessing .QueryTable
    If lo.SourceType = xlSrcRange Then Exit Function

    On Error GoTo NotAQueryTable
    Dim qt As QueryTable
    Set qt = lo.QueryTable
    If qt.WorkbookConnection Is conn Then
        GetFastDataLoad = Not qt.PreserveFormatting
    End If
    Exit Function
NotAQueryTable:
    GetFastDataLoad = Empty
End Function
