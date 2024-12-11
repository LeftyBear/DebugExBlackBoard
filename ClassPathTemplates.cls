VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPath
    Set ParentPath = New ClassPath
End Function
Public Function FileDaily() As String
    FileDaily = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Daily.xlsx")
End Function
Public Function FileDailyTemplateSetting() As String
    FileDailyTemplateSetting = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Daily.csv")
End Function
Public Function FileDailyManual() As String
    FileDailyManual = FSO.BuildPath(Me.ParentPath.FolderTemplates, "DailyManual.xlsx")
End Function
Public Function FileMonthly() As String
    FileMonthly = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Monthly.xlsx")
End Function
Public Function FileMonthlyTemplateSetting() As String
    FileMonthlyTemplateSetting = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Monthly.csv")
End Function
Public Function FileWeekly() As String
    FileWeekly = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Weekly.xlsx")
End Function
Public Function FileWeeklyTemplateSetting() As String
    FileWeeklyTemplateSetting = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Weekly.csv")
End Function
Public Function FileYearly() As String
    FileYearly = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Yearly.xlsx")
End Function
Public Function FileYearlyTemplateSetting() As String
    FileYearlyTemplateSetting = FSO.BuildPath(Me.ParentPath.FolderTemplates, "Yearly.csv")
End Function
