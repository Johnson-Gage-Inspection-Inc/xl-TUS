VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Calculate()
    ' When UUTRangeTol recalculates (e.g. order number change cascades
    ' through WorkItemDetails -> LET/FILTER -> Unit formula), check whether
    ' the Unit value now differs from the current number formats and swap
    ' if needed.  ApplyUnitFormats probes Main!D15.NumberFormat and exits
    ' immediately when formats already match, so the cost is negligible.
    ApplyUnitFormats
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
