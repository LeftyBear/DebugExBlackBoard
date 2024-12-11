VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathRegular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPathData
    Set ParentPath = New ClassPathData
End Function
Public Function FileRegularDays() As String
    FileRegularDays = FSO.BuildPath(Me.ParentPath.FolderRegular, "RegularDays.csv")
End Function
Public Function FileRegularWeekDays() As String
    FileRegularWeekDays = FSO.BuildPath(Me.ParentPath.FolderRegular, "RegularWeekDays.csv")
End Function

