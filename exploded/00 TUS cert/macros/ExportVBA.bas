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
    exportPath = ActiveWorkbook.path & "\exploded\" & fileBaseName & "\macros"
    count = 0

    If Not fso.FolderExists(exportPath) Then
        Call fso.CreateFolder(exportPath)
    End If

    totalComponents = ActiveWorkbook.VBProject.VBComponents.count
    progress = 0

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
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

