VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPathData
    Set ParentPath = New ClassPathData
End Function
Public Function FolderByYear(ByVal WesternYear As String) As String
    FolderByYear = FSO.BuildPath(Me.ParentPath.FolderSchedules, WesternYear) & AS_FOLDER
End Function
Public Function FileByYear(ByVal WesternYear As String) As String
    FileByYear = FSO.BuildPath(FolderByYear(WesternYear), "schedule" & WesternYear & ".csv")
End Function
Public Function FileHeader() As String
    FileHeader = FSO.BuildPath(Me.ParentPath.FolderSchedules, "ScheduleHeader.csv")
End Function

