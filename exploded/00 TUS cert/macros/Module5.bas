Attribute VB_Name = "Module5"
Sub Update_Header_Info()
Attribute Update_Header_Info.VB_Description = "Update Header Info"
Attribute Update_Header_Info.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Update_Header_Info Macro
' Update Header Info
'

'
    Range("A2:AG700").Select
    Selection.Copy
    Windows("TUS cert 02-25-20-01.xlsm").Activate
    Sheets("Header Info").Select
    ActiveWorkbook.Names.Add Name:="Locator", RefersToR1C1:= _
        "='Header Info'!R524288C8192"
    Range("A2").Select
    Application.WindowState = xlNormal
    Windows("Header Info.xlsx").Activate
    Selection.Copy
    Windows("TUS cert 02-25-20-01.xlsm").Activate
    ActiveSheet.Paste
    Sheets("Main").Select
    ActiveWorkbook.Names.Add Name:="Locator", RefersToR1C1:= _
        "=Main!R524288C8192"
    Range("D3").Select
End Sub
