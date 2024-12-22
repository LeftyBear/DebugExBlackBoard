VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BootObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Boot")
Option Explicit
Implements IBootObject
Private AddinSheetHoldsItem As IAddinSheetHoldsItem
Private Sub Class_Initialize()
    Set AddinSheetHoldsItem = New AddinSheetHoldsItem
End Sub
Private Function IBootObject_TargetBook() As Workbook
    Set IBootObject_TargetBook = Workbooks.Item(AddinSheetHoldsItem.BootBookName)
End Function
Private Function IBootObject_TargetSheet() As Worksheet
    Dim BootBook As Workbook
    Set BootBook = Workbooks.Item(AddinSheetHoldsItem.BootBookName)
    Set IBootObject_TargetSheet = BootBook.Worksheets.Item(AddinSheetHoldsItem.BootSheetName)
End Function
