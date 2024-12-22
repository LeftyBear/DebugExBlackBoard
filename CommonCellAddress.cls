VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonCellAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CommonCellAddress")
Option Explicit
Implements ICommonCellAddress
Private Const ADDRESS_ADDIN_NAME                    As String = "B1"
Private Const ADDRESS_ADDIN_VERSION                 As String = "B2"
Private Const ADDRESS_TOP_FOLDER_PATH               As String = "B3"
Private Const ADDRESS_BOOT_BOOK_NAME                As String = "B4"
Private Const ADDRESS_BOOT_SHEET_NAME               As String = "B5"
Private Const ADDRESS_CURRENT_YEAR                  As String = "B6"
Private Const ADDRESS_SCHEDULE_UPDATEAT             As String = "B7"
Private Const ADDRESS_LOCAL_BACKUP_FOLDER_PATH      As String = "B8"
Private Function ICommonCellAddress_ForAddinName() As String
    ICommonCellAddress_ForAddinName = ADDRESS_ADDIN_NAME
End Function
Private Function ICommonCellAddress_ForAddinVersion() As String
    ICommonCellAddress_ForAddinVersion = ADDRESS_ADDIN_VERSION
End Function
Private Function ICommonCellAddress_ForBootBookName() As String
    ICommonCellAddress_ForBootBookName = ADDRESS_BOOT_BOOK_NAME
End Function
Private Function ICommonCellAddress_ForBootSheetName() As String
    ICommonCellAddress_ForBootSheetName = ADDRESS_BOOT_SHEET_NAME
End Function
Private Function ICommonCellAddress_ForLocalBackupFolderPath() As String
    ICommonCellAddress_ForLocalBackupFolderPath = ADDRESS_LOCAL_BACKUP_FOLDER_PATH
End Function
Private Function ICommonCellAddress_ForCurrentYear() As String
    ICommonCellAddress_ForCurrentYear = ADDRESS_CURRENT_YEAR
End Function
Private Function ICommonCellAddress_ForScheduleUpDateAt() As String
    ICommonCellAddress_ForScheduleUpDateAt = ADDRESS_SCHEDULE_UPDATEAT
End Function
Private Function ICommonCellAddress_ForTopFolderPath() As String
    ICommonCellAddress_ForTopFolderPath = ADDRESS_TOP_FOLDER_PATH
End Function
