Attribute VB_Name = "ExportVBA"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
' Requires check "Microsoft Scripting Runtime" at Tools > References to use "FileSystemObject";

Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Long
    Dim path As String
    Dim exportPath As String
    Dim extension As String
    Dim fso As Object
    Dim totalComponents As Long
    Dim progress As Long
    Dim fileBaseName As String
    fileBaseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    exportPath = ThisWorkbook.path & "\exploded\" & fileBaseName & "\macros"
    count = 0

    If Not fso.FolderExists(exportPath) Then
        Call fso.CreateFolder(exportPath)
    End If

    totalComponents = ThisWorkbook.VBProject.VBComponents.count
    progress = 0

    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        progress = progress + 1
        Application.StatusBar = "Exporting VBA (" & progress & " of " & totalComponents & "): " & VBComponent.Name

        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select

        On Error Resume Next
        Err.Clear

        path = exportPath & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)

        If Err.Number <> 0 Then
            MsgBox "Failed to export " & VBComponent.Name & " to " & path, vbCritical
        Else
            Err.Clear
            Call ConvertAnsiFileToUTF8(path)
            If Err.Number <> 0 Then
                MsgBox "Exported but UTF-8 conversion failed for " & VBComponent.Name & ": " & Err.Description, vbExclamation
                Err.Clear
            End If
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next

    Application.StatusBar = "? Successfully exported " & CStr(count) & " VBA files to " & exportPath
    Application.OnTime Now + TimeSerial(0, 0, 5), "ClearStatusBar"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' Re-encode a file from the system ANSI codepage to UTF-8 without BOM.
' VBComponent.Export always writes ANSI; this post-processes each file so git
' (and every other modern tool) sees valid UTF-8.
' Errors are raised to the caller via Err.Raise so they are not silently lost.
Private Sub ConvertAnsiFileToUTF8(filePath As String)
    Dim txtStream As Object
    Dim binStream As Object
    Dim fso As Object
    Dim tempPath As String
    Dim content As String
    Dim utf8Bytes() As Byte

    On Error GoTo ConvertFailed

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = filePath & ".tmp"

    ' 1. Read the ANSI file using the system's active codepage
    Set txtStream = CreateObject("ADODB.Stream")
    txtStream.Type = 2            ' adTypeText
    txtStream.Charset = "x-ansi"
    txtStream.Open
    txtStream.LoadFromFile filePath
    content = txtStream.ReadText(-1)  ' adReadAll
    txtStream.Close

    ' 2. Write the content as UTF-8 (ADODB adds a 3-byte BOM)
    txtStream.Charset = "UTF-8"
    txtStream.Open
    txtStream.WriteText content
    txtStream.Position = 0
    txtStream.Type = 1            ' switch to adTypeBinary while at Position 0
    txtStream.Position = 3        ' skip the BOM (EF BB BF)

    utf8Bytes = txtStream.Read(-1)    ' remaining bytes = UTF-8 without BOM
    txtStream.Close

    ' 3. Save BOM-free bytes to a temp file first
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1            ' adTypeBinary
    binStream.Open
    binStream.Write utf8Bytes
    binStream.SaveToFile tempPath, 2  ' adSaveCreateOverWrite
    binStream.Close

    ' 4. Replace the original only after a successful write
    If fso.FileExists(tempPath) Then
        fso.DeleteFile filePath
        fso.MoveFile tempPath, filePath
    End If

    Set txtStream = Nothing
    Set binStream = Nothing
    Set fso = Nothing
    Exit Sub

ConvertFailed:
    ' Clean up the temp file if it was partially written
    On Error Resume Next
    If Not txtStream Is Nothing Then txtStream.Close
    If Not binStream Is Nothing Then binStream.Close
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    End If
    Set txtStream = Nothing
    Set binStream = Nothing
    Set fso = Nothing
    On Error GoTo 0
    Err.Raise vbObjectError + 1, "ConvertAnsiFileToUTF8", _
        "UTF-8 conversion failed for " & filePath & ": " & Err.Description
End Sub

