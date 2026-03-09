Attribute VB_Name = "ExportPowerQuery"
Option Explicit

Sub ExportAllQueryMCode()
    Dim q As WorkbookQuery
    Dim regEx As Object
    Dim exportPath As String, mCode As String, fileBaseName As String
    fileBaseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    exportPath = ThisWorkbook.path & "\exploded\" & fileBaseName & "\queries"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.pattern = "Authorization\s*=\s*""Api-Token [^""]+"""

    Dim connHeader As String
    Dim conn As WorkbookConnection
    Dim oledb As OLEDBConnection

    ' Pre-compute EnableFastDataLoad map from DataMashup metadata
    Dim fdlMap As Object
    Set fdlMap = GetFastDataLoadMap()

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
            If fdlMap.Exists(q.Name) Then
                connHeader = connHeader & "//   EnableFastDataLoad:    " & fdlMap(q.Name) & vbLf
            End If
            connHeader = connHeader & vbLf
        End If

        ' Write as UTF-8 (no BOM) via ADODB.Stream
        Dim stm As Object
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2          ' adTypeText
        stm.Charset = "utf-8"
        stm.Open
        stm.WriteText connHeader & mCode

        ' ADODB.Stream prepends a 3-byte BOM; strip it
        stm.Position = 0
        stm.Type = 1          ' adTypeBinary
        stm.Position = 3      ' skip BOM
        Dim fileBody() As Byte
        fileBody = stm.Read
        stm.Close

        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 1
        stm.Open
        stm.Write fileBody
        stm.SaveToFile exportPath & "\" & q.Name & ".m", 2  ' adSaveCreateOverWrite
        stm.Close

        Debug.Print q.Name & " exported"
    Next q
End Sub

Private Function GetFastDataLoadMap() As Object
    ' Returns a Dictionary mapping query name -> Boolean (EnableFastDataLoad).
    '
    ' EnableFastDataLoad is stored as the inverse of "BufferNextRefresh" in the
    ' LocalPackageMetadataFile embedded inside the DataMashup binary blob
    ' (stored in a CustomXMLPart).
    '
    ' Binary layout of the DataMashup blob:
    '   version(4) + pkg_len(4) + ZIP(pkg_len)
    '   + perm_len(4) + perm(perm_len)
    '   + meta_len(4) + meta(meta_len)
    '
    ' The metadata block has its own sub-header:
    '   sub_version(4) + xml_length(4) + xml(xml_length)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    On Error GoTo Done

    ' --- Find the DataMashup CustomXMLPart ---
    Dim cxp As CustomXMLPart
    Dim mashupXml As String
    Dim found As Boolean: found = False

    For Each cxp In ThisWorkbook.CustomXMLParts
        If InStr(1, cxp.XML, "DataMashup", vbTextCompare) > 0 Then
            mashupXml = cxp.XML
            found = True
            Exit For
        End If
    Next cxp

    If Not found Then GoTo Done

    ' --- Decode the base-64 payload ---
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML mashupXml

    Dim b64Text As String
    b64Text = Trim$(xmlDoc.DocumentElement.Text)
    If Len(b64Text) = 0 Then GoTo Done

    Dim b64Node As Object
    Set b64Node = xmlDoc.createElement("b64")
    b64Node.DataType = "bin.base64"
    b64Node.Text = b64Text
    Dim raw() As Byte
    raw = b64Node.nodeTypedValue

    ' --- Navigate the binary structure ---
    Dim pos As Long: pos = 4                              ' skip version
    Dim pkgLen As Long: pkgLen = ReadUInt32LE(raw, pos)
    pos = pos + 4 + pkgLen                                ' skip ZIP package

    Dim permLen As Long: permLen = ReadUInt32LE(raw, pos)
    pos = pos + 4 + permLen                               ' skip permissions

    Dim metaLen As Long: metaLen = ReadUInt32LE(raw, pos)
    pos = pos + 4                                         ' now at metadata block

    ' Metadata sub-header: sub_version(4) + xml_length(4)
    Dim xmlLen As Long: xmlLen = ReadUInt32LE(raw, pos + 4)

    ' Copy metadata XML bytes so we can decode UTF-8
    Dim metaBytes() As Byte
    ReDim metaBytes(xmlLen - 1)
    Dim i As Long
    For i = 0 To xmlLen - 1
        metaBytes(i) = raw(pos + 8 + i)
    Next i

    ' UTF-8 -> VBA Unicode string via ADODB.Stream
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1  ' adTypeBinary
    stm.Open
    stm.Write metaBytes
    stm.Position = 0
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    Dim metaXmlStr As String
    metaXmlStr = stm.ReadText
    stm.Close

    ' Strip BOM if present
    If Len(metaXmlStr) > 0 Then
        If AscW(Left$(metaXmlStr, 1)) = &HFEFF Then
            metaXmlStr = Mid$(metaXmlStr, 2)
        End If
    End If

    ' --- Parse the metadata XML ---
    Dim metaDoc As Object
    Set metaDoc = CreateObject("MSXML2.DOMDocument")
    If Not metaDoc.LoadXML(metaXmlStr) Then GoTo Done

    Dim itemsNode As Object
    Set itemsNode = metaDoc.DocumentElement.ChildNodes(0)  ' <Items>

    Dim j As Long, k As Long, m As Long, n As Long
    For j = 0 To itemsNode.ChildNodes.Length - 1
        Dim itemNode As Object
        Set itemNode = itemsNode.ChildNodes(j)

        ' Find ItemLocation / ItemPath -> query name
        Dim queryName As String: queryName = ""
        For k = 0 To itemNode.ChildNodes.Length - 1
            If itemNode.ChildNodes(k).BaseName = "ItemLocation" Then
                Dim locNode As Object
                Set locNode = itemNode.ChildNodes(k)
                For m = 0 To locNode.ChildNodes.Length - 1
                    If locNode.ChildNodes(m).BaseName = "ItemPath" Then
                        Dim pathText As String
                        pathText = Trim$(locNode.ChildNodes(m).Text)
                        Dim slashPos As Long
                        slashPos = InStrRev(pathText, "/")
                        If slashPos > 0 Then
                            queryName = Mid$(pathText, slashPos + 1)
                        Else
                            queryName = pathText
                        End If
                        Exit For
                    End If
                Next m
                Exit For
            End If
        Next k

        If Len(queryName) = 0 Then GoTo NextItem

        ' Find StableEntries -> BufferNextRefresh
        For k = 0 To itemNode.ChildNodes.Length - 1
            If itemNode.ChildNodes(k).BaseName = "StableEntries" Then
                Dim stableNode As Object
                Set stableNode = itemNode.ChildNodes(k)
                For n = 0 To stableNode.ChildNodes.Length - 1
                    Dim entryNode As Object
                    Set entryNode = stableNode.ChildNodes(n)
                    If entryNode.getAttribute("Type") = "BufferNextRefresh" Then
                        ' "l0" -> FastDataLoad = True;  "l1" -> FastDataLoad = False
                        dict(queryName) = (entryNode.getAttribute("Value") = "l0")
                        Exit For
                    End If
                Next n
                Exit For
            End If
        Next k
NextItem:
    Next j

Done:
    Set GetFastDataLoadMap = dict
End Function

Private Function ReadUInt32LE(buf() As Byte, ByVal pos As Long) As Long
    ' Read a little-endian unsigned 32-bit integer from buf at position pos.
    ' Safe for values up to 2^31-1 (more than enough for DataMashup segments).
    ReadUInt32LE = CLng(buf(pos)) _
                 + CLng(buf(pos + 1)) * &H100& _
                 + CLng(buf(pos + 2)) * &H10000 _
                 + CLng(buf(pos + 3)) * &H1000000
End Function

