VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub UpdateClientAssets(sCustomerName As String)
    If Trim$(sCustomerName) = "" Then
        ThisWorkbook.Worksheets("Main").Range("D6:H6").ClearContents
        Exit Sub
    End If

    Me.Unprotect Password:="JGIPyro"

    ThisWorkbook.Connections("Query - ClientAssets_GetAssetsByClientCompanyId").Refresh

    Application.OnTime Now + TimeValue("00:00:02"), "FinalizeClientAssetsProtection"
End Sub
