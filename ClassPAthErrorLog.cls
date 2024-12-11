VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPAthErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPathData
    Set ParentPath = New ClassPathData
End Function
Public Function FileErrorLog() As String
    FileErrorLog = FSO.BuildPath(Me.ParentPath.FolderErrorLog, "error.log")
End Function

