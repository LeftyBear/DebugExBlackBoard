VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPath
    Set ParentPath = New ClassPath
End Function
Public Function cache() As ClassPathCache
    Set cache = New ClassPathCache
End Function
Public Function ClassHours() As ClassPathClassHours
    Set ClassHours = New ClassPathClassHours
End Function
Public Function dayduty() As ClassPathDayDuty
    Set dayduty = New ClassPathDayDuty
End Function
Public Function enrollments() As ClassPathEnrollments
    Set enrollments = New ClassPathEnrollments
End Function
Public Function errorlog() As ClassPAthErrorLog
    Set errorlog = New ClassPAthErrorLog
End Function
Public Function Holiday() As ClassPathHoliday
    Set Holiday = New ClassPathHoliday
End Function
Public Function leave() As ClassPathLeave
    Set leave = New ClassPathLeave
End Function
Public Function regular() As ClassPathRegular
    Set regular = New ClassPathRegular
End Function
Public Function schedules() As ClassPathSchedules
    Set schedules = New ClassPathSchedules
End Function
Public Function FolderBackup() As String
    FolderBackup = FSO.BuildPath(Me.ParentPath.FolderData, "backup") & AS_FOLDER
End Function
Public Function FolderCache() As String
    FolderCache = FSO.BuildPath(Me.ParentPath.FolderData, "cache") & AS_FOLDER
End Function
Public Function FolderClasshours() As String
    FolderClasshours = FSO.BuildPath(Me.ParentPath.FolderData, "classhours") & AS_FOLDER
End Function
Public Function FolderDayDuty() As String
    FolderDayDuty = FSO.BuildPath(Me.ParentPath.FolderData, "dayduty") & AS_FOLDER
End Function
Public Function FolderEnrollments() As String
    FolderEnrollments = FSO.BuildPath(Me.ParentPath.FolderData, "enrollments") & AS_FOLDER
End Function
Public Function FolderErrorLog() As String
    FolderErrorLog = FSO.BuildPath(Me.ParentPath.FolderData, "errorlog") & AS_FOLDER
End Function
Public Function FolderHoliday() As String
    FolderHoliday = FSO.BuildPath(Me.ParentPath.FolderData, "holiday") & AS_FOLDER
End Function
Public Function FolderLeave() As String
    FolderLeave = FSO.BuildPath(Me.ParentPath.FolderData, "leave") & AS_FOLDER
End Function
Public Function FolderRegular() As String
    FolderRegular = FSO.BuildPath(Me.ParentPath.FolderData, "regular") & AS_FOLDER
End Function
Public Function FolderSchedules() As String
    FolderSchedules = FSO.BuildPath(Me.ParentPath.FolderData, "schedules") & AS_FOLDER
End Function
Public Function FolderWriteLog() As String
    FolderWriteLog = FSO.BuildPath(Me.ParentPath.FolderData, "writelog") & AS_FOLDER
End Function
