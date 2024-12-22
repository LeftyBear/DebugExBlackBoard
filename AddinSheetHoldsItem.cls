VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddinSheetHoldsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Addin")
Option Explicit
Implements IAddinSheetHoldsItem
Private TargetSheet         As Worksheet
Private CommonCellAddress   As ICommonCellAddress
Private Sub Class_Initialize()
    Set TargetSheet = AddinSheet
    Set CommonCellAddress = New CommonCellAddress
End Sub
Private Property Let IAddinSheetHoldsItem_BootBookName(ByVal BootBookName As String)
    TargetSheet.Range(CommonCellAddress.ForBootBookName).Value = BootBookName
End Property
Private Property Get IAddinSheetHoldsItem_BootBookName() As String
    IAddinSheetHoldsItem_BootBookName = TargetSheet.Range(CommonCellAddress.ForBootBookName).Value
End Property
Private Property Let IAddinSheetHoldsItem_BootSheetName(ByVal BootSheetName As String)
    TargetSheet.Range(CommonCellAddress.ForBootSheetName).Value = BootSheetName
End Property
Private Property Get IAddinSheetHoldsItem_BootSheetName() As String
    IAddinSheetHoldsItem_BootSheetName = TargetSheet.Range(CommonCellAddress.ForBootSheetName).Value
End Property
