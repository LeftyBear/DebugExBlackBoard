VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BootSheetHoldsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MoveFieldCloserToUsage
'@Folder("Boot")
Option Explicit
Implements IBootSheetHoldsItem
Private CommonCellAddress   As ICommonCellAddress
Private BootObject          As IBootObject
Private TargetSheet         As Worksheet
Private Sub Class_Initialize()
    Set CommonCellAddress = New CommonCellAddress
    Set BootObject = New BootObject
    Set TargetSheet = BootObject.TargetSheet
End Sub
Private Property Let IBootSheetHoldsItem_AddinName(ByVal AddinName As String)
    TargetSheet.Range(CommonCellAddress.ForAddinName).Value = AddinName
End Property
Private Property Get IBootSheetHoldsItem_AddinName() As String
    IBootSheetHoldsItem_AddinName = TargetSheet.Range(CommonCellAddress.ForAddinName).Value
End Property
Private Property Let IBootSheetHoldsItem_AddinVersion(ByVal AddinVersion As String)
    TargetSheet.Range(CommonCellAddress.ForAddinVersion).Value = AddinVersion
End Property
Private Property Get IBootSheetHoldsItem_AddinVersion() As String
    IBootSheetHoldsItem_AddinVersion = TargetSheet.Range(CommonCellAddress.ForAddinVersion).Value
End Property
Private Property Let IBootSheetHoldsItem_CurrentScheduleFileUpdateAt(ByVal UpdateAt As Date)
    TargetSheet.Range(CommonCellAddress.ForScheduleUpDateAt).Value = UpdateAt
End Property
Private Property Get IBootSheetHoldsItem_CurrentScheduleFileUpdateAt() As Date
    IBootSheetHoldsItem_CurrentScheduleFileUpdateAt = TargetSheet.Range(CommonCellAddress.ForScheduleUpDateAt).Value
End Property
Private Property Let IBootSheetHoldsItem_CurrentYear(ByVal WesternYear As String)
    TargetSheet.Range(CommonCellAddress.ForCurrentYear).Value = WesternYear
End Property
Private Property Get IBootSheetHoldsItem_CurrentYear() As String
    IBootSheetHoldsItem_CurrentYear = TargetSheet.Range(CommonCellAddress.ForCurrentYear).Value
End Property
Private Property Let IBootSheetHoldsItem_LocalBackupFolderPath(ByVal FolderPath As String)
    TargetSheet.Range(CommonCellAddress.ForLocalBackupFolderPath).Value = FolderPath
End Property
Private Property Get IBootSheetHoldsItem_LocalBackupFolderPath() As String
    IBootSheetHoldsItem_LocalBackupFolderPath = TargetSheet.Range(CommonCellAddress.ForLocalBackupFolderPath).Value
End Property
Private Function IBootSheetHoldsItem_TopFolderPath() As String
    IBootSheetHoldsItem_TopFolderPath = TargetSheet.Range(CommonCellAddress.ForTopFolderPath).Value
End Function
