VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISystemPath
Private topFolder As String
Private Property Let ISystemPath_topFolder(ByVal FolderPath As String)
    topFolder = FolderPath
End Property
Private Property Get ISystemPath_topFolder() As String
    ISystemPath_topFolder = topFolder
End Property
Private Function ISystemPath_rootFolder() As String
    ISystemPath_rootFolder = FSO.BuildPath(ISystemPath_topFolder, "root") & AS_FOLDER
End Function
Private Function ISystemPath_dataFolder() As String
    ISystemPath_dataFolder = FSO.BuildPath(ISystemPath_rootFolder, "data") & AS_FOLDER
End Function
Private Function ISystemPath_backupFolder() As String
    ISystemPath_backupFolder = FSO.BuildPath(ISystemPath_dataFolder, "backup") & AS_FOLDER
End Function
Private Function ISystemPath_cacheFolder() As String
    ISystemPath_cacheFolder = FSO.BuildPath(ISystemPath_dataFolder, "cache") & AS_FOLDER
End Function
Private Function ISystemPath_AdministerFile() As String
    ISystemPath_AdministerFile = FSO.BuildPath(ISystemPath_cacheFolder, "Administer.txt")
End Function
Private Function ISystemPath_OnTimeFile() As String
    ISystemPath_OnTimeFile = FSO.BuildPath(ISystemPath_cacheFolder, "OnTime.log")
End Function
Private Function ISystemPath_PasswordFile() As String
    ISystemPath_PasswordFile = FSO.BuildPath(ISystemPath_cacheFolder, "Password.txt")
End Function
Private Function ISystemPath_PasswordSkipFile() As String
    ISystemPath_PasswordSkipFile = FSO.BuildPath(ISystemPath_cacheFolder, "PasswordSkip.txt")
End Function
Private Function ISystemPath_classhoursFolder() As String
    ISystemPath_classhoursFolder = FSO.BuildPath(ISystemPath_dataFolder, "classhours") & AS_FOLDER
End Function
Private Function ISystemPath_classhoursFolderByYear(ByVal WesternYear As String) As String
    ISystemPath_classhoursFolderByYear = FSO.BuildPath(ISystemPath_classhoursFolder, WesternYear) & AS_FOLDER
End Function
Private Function ISystemPath_ClasshoursFileByYear(ByVal WesternYear As String) As String
    ISystemPath_ClasshoursFileByYear = FSO.BuildPath(ISystemPath_classhoursFolderByYear(WesternYear), "classhour" & WesternYear & ".csv")
End Function
Private Function ISystemPath_daydutyFolder() As String
    ISystemPath_daydutyFolder = FSO.BuildPath(ISystemPath_dataFolder, "dayduty") & AS_FOLDER
End Function
Private Function ISystemPath_DayDutyFile() As String
    ISystemPath_DayDutyFile = FSO.BuildPath(ISystemPath_daydutyFolder, "dayduty.csv")
End Function
Private Function ISystemPath_enrollmentsFolder() As String
    ISystemPath_enrollmentsFolder = FSO.BuildPath(ISystemPath_dataFolder, "enrollments") & AS_FOLDER
End Function
Private Function ISystemPath_enrollmentsFolderByYear(ByVal WesternYear As String) As String
    ISystemPath_enrollmentsFolderByYear = FSO.BuildPath(ISystemPath_enrollmentsFolder, WesternYear) & AS_FOLDER
End Function
Private Function ISystemPath_EnrollmentsFileByYear(ByVal WesternYear As String) As String
    ISystemPath_EnrollmentsFileByYear = FSO.BuildPath(ISystemPath_enrollmentsFolderByYear(WesternYear), "enrollment" & WesternYear & ".csv")
End Function
Private Function ISystemPath_errorlogFolder() As String
    ISystemPath_errorlogFolder = FSO.BuildPath(ISystemPath_dataFolder, "errorlog") & AS_FOLDER
End Function
Private Function ISystemPath_leaveFolder() As String
    ISystemPath_leaveFolder = FSO.BuildPath(ISystemPath_dataFolder, "leave") & AS_FOLDER
End Function
Private Function ISystemPath_regularFolder() As String
    ISystemPath_regularFolder = FSO.BuildPath(ISystemPath_dataFolder, "regular") & AS_FOLDER
End Function
Private Function ISystemPath_RegularDaysFile() As String
    ISystemPath_RegularDaysFile = FSO.BuildPath(ISystemPath_regularFolder, "RegularDays.csv")
End Function
Private Function ISystemPath_RegularWeekDaysFile() As String
    ISystemPath_RegularWeekDaysFile = FSO.BuildPath(ISystemPath_regularFolder, "RegularWeekDays.csv")
End Function
Private Function ISystemPath_schedulesFolder() As String
    ISystemPath_schedulesFolder = FSO.BuildPath(ISystemPath_dataFolder, "schedules") & AS_FOLDER
End Function
Private Function ISystemPath_schedulesFolderByYear(ByVal WesternYear As String) As String
    ISystemPath_schedulesFolderByYear = FSO.BuildPath(ISystemPath_schedulesFolder, WesternYear) & AS_FOLDER
End Function
Private Function ISystemPath_SchedulesFileByYear(ByVal WesternYear As String) As String
    ISystemPath_SchedulesFileByYear = FSO.BuildPath(ISystemPath_schedulesFolderByYear(WesternYear), "schedule" & WesternYear & ".csv")
End Function
Private Function ISystemPath_writelogFolder() As String
    ISystemPath_writelogFolder = FSO.BuildPath(ISystemPath_dataFolder, "writelog") & AS_FOLDER
End Function
Private Function ISystemPath_headersFolder() As String
    ISystemPath_headersFolder = FSO.BuildPath(ISystemPath_rootFolder, "headers") & AS_FOLDER
End Function
Private Function ISystemPath_ClassHourHeaderFile() As String
    ISystemPath_ClassHourHeaderFile = FSO.BuildPath(ISystemPath_headersFolder, "ClassHourHeader.csv")
End Function
Private Function ISystemPath_EnrollmentHeaderFile() As String
    ISystemPath_EnrollmentHeaderFile = FSO.BuildPath(ISystemPath_headersFolder, "EnrollmentHeader.csv")
End Function
Private Function ISystemPath_ScheduleHeaderFile() As String
    ISystemPath_ScheduleHeaderFile = FSO.BuildPath(ISystemPath_headersFolder, "ScheduleHeader.csv")
End Function
Private Function ISystemPath_holidayFolder() As String
    ISystemPath_holidayFolder = FSO.BuildPath(ISystemPath_rootFolder, "holiday") & AS_FOLDER
End Function
Private Function ISystemPath_HolidayFile() As String
    ISystemPath_HolidayFile = FSO.BuildPath(ISystemPath_holidayFolder, "Holiday.csv")
End Function
Private Function ISystemPath_icoFolder() As String
    ISystemPath_icoFolder = FSO.BuildPath(ISystemPath_topFolder, "ico") & AS_FOLDER
End Function
Private Function ISystemPath_CursorFile() As String
    ISystemPath_CursorFile = FSO.BuildPath(ISystemPath_icoFolder, "cursor.csv")
End Function
Private Function ISystemPath_srcFolder() As String
    ISystemPath_srcFolder = FSO.BuildPath(ISystemPath_topFolder, "src") & AS_FOLDER
End Function
Private Function ISystemPath_templatesFolder() As String
    ISystemPath_templatesFolder = FSO.BuildPath(ISystemPath_topFolder, "templates") & AS_FOLDER
End Function
Private Function ISystemPath_DailyFile() As String
    ISystemPath_DailyFile = FSO.BuildPath(ISystemPath_templatesFolder, "Daily.xlsx")
End Function
Private Function ISystemPath_DailyManualFile() As String
    ISystemPath_DailyManualFile = FSO.BuildPath(ISystemPath_templatesFolder, "DailyManual.xlsx")
End Function
Private Function ISystemPath_MonthlyFile() As String
    ISystemPath_MonthlyFile = FSO.BuildPath(ISystemPath_templatesFolder, "Monthly.xlsx")
End Function
Private Function ISystemPath_WeeklyFile() As String
    ISystemPath_WeeklyFile = FSO.BuildPath(ISystemPath_templatesFolder, "Weekly.xlsx")
End Function
Private Function ISystemPath_YearlyFile() As String
    ISystemPath_YearlyFile = FSO.BuildPath(ISystemPath_templatesFolder, "Yearly.xlsx")
End Function

