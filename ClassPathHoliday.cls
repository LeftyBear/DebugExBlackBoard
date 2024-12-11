VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPathData
    Set ParentPath = New ClassPathData
End Function
Public Function FileHoliday() As String
    FileHoliday = FSO.BuildPath(Me.ParentPath.FolderHoliday, "syukujitsu.csv")
End Function

