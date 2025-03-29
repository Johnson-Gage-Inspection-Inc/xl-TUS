Attribute VB_Name = "Module6"
Sub Create_Customer_List()
Attribute Create_Customer_List.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_Customer_List Macro
'

'
    Range("B2:B546").Select
    Selection.Copy
    Range("CM2").Select
    ActiveWindow.SmallScroll Down:=-1
    Range("CG2").Select
    ActiveSheet.Paste
    ActiveSheet.Range("$CG$1:$CG$546").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("B1").Select
End Sub
