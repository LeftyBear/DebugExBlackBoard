VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBootSheetHoldsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule VariableNotAssigned, VariableNotUsed, MoveFieldCloserToUsage, EncapsulatePublicField
'@Folder("Boot")
Option Explicit
Public AddinName                    As String
Public AddinVersion                 As String
Public CurrentYear                  As String
Public CurrentScheduleFileUpdateAt  As Date
Public LocalBackupFolderPath        As String
Public Function TopFolderPath() As String
End Function
