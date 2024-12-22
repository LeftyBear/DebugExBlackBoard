VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed, MoveFieldCloserToUsage
'@Folder("Directory")
Option Explicit
Implements IDirectory
Private BootSheetHoldsItem As IBootSheetHoldsItem
Private TopFolderPath As String
Private Sub Class_Initialize()
    Set BootSheetHoldsItem = New BootSheetHoldsItem
    IDirectory_TopFolderPath = BootSheetHoldsItem.TopFolderPath
End Sub
Private Property Let IDirectory_TopFolderPath(ByVal FolderPath As String)
    TopFolderPath = FolderPath
End Property
Private Property Get IDirectory_TopFolderPath() As String
    IDirectory_TopFolderPath = TopFolderPath
End Property
Private Function IDirectory_topFolder() As String
        IDirectory_topFolder = IDirectory_TopFolderPath
End Function
Private Function IDirectory_rootFolder() As String
    IDirectory_rootFolder = FSO.BuildPath(IDirectory_topFolder, "root") & AS_FOLDER
End Function
Private Function IDirectory_dataFolder() As String
    IDirectory_dataFolder = FSO.BuildPath(IDirectory_rootFolder, "data") & AS_FOLDER
End Function
Private Function IDirectory_backupFolder() As String
    IDirectory_backupFolder = FSO.BuildPath(IDirectory_dataFolder, "backup") & AS_FOLDER
End Function
Private Function IDirectory_cacheFolder() As String
    IDirectory_cacheFolder = FSO.BuildPath(IDirectory_dataFolder, "cache") & AS_FOLDER
End Function
Private Function IDirectory_AdministorFile() As String
    IDirectory_AdministorFile = FSO.BuildPath(IDirectory_cacheFolder, "Administrator.txt")
End Function
Private Function IDirectory_SwappingCheckFile() As String
    IDirectory_SwappingCheckFile = FSO.BuildPath(IDirectory_cacheFolder, "OnTime.log")
End Function
Private Function IDirectory_PasswordFile() As String
    IDirectory_PasswordFile = FSO.BuildPath(IDirectory_cacheFolder, "Password.txt")
End Function
Private Function IDirectory_PasswordSkipFile() As String
    IDirectory_PasswordSkipFile = FSO.BuildPath(IDirectory_cacheFolder, "PasswordSkip.txt")
End Function
Private Function IDirectory_classhoursFolder() As String
    IDirectory_classhoursFolder = FSO.BuildPath(IDirectory_dataFolder, "classhours") & AS_FOLDER
End Function
Private Function IDirectory_classhoursFolderByYear(ByVal WesternYear As String) As String
    IDirectory_classhoursFolderByYear = FSO.BuildPath(IDirectory_classhoursFolder, WesternYear) & AS_FOLDER
End Function
Private Function IDirectory_ClasshourFileByYear(ByVal WesternYear As String) As String
    IDirectory_ClasshourFileByYear = FSO.BuildPath(IDirectory_classhoursFolderByYear(WesternYear), "classhour" & WesternYear & ".csv")
End Function
Private Function IDirectory_daydutyFolder() As String
    IDirectory_daydutyFolder = FSO.BuildPath(IDirectory_dataFolder, "dayduty") & AS_FOLDER
End Function
Private Function IDirectory_DayDutyFile() As String
    IDirectory_DayDutyFile = FSO.BuildPath(IDirectory_daydutyFolder, "dayduty.csv")
End Function
Private Function IDirectory_enrollmentsFolder() As String
    IDirectory_enrollmentsFolder = FSO.BuildPath(IDirectory_dataFolder, "enrollments") & AS_FOLDER
End Function
Private Function IDirectory_enrollmentsFolderByYear(ByVal WesternYear As String) As String
    IDirectory_enrollmentsFolderByYear = FSO.BuildPath(IDirectory_enrollmentsFolder, WesternYear) & AS_FOLDER
End Function
Private Function IDirectory_EnrollmentFileByYear(ByVal WesternYear As String) As String
    IDirectory_EnrollmentFileByYear = FSO.BuildPath(IDirectory_enrollmentsFolderByYear(WesternYear), "enrollment" & WesternYear & ".csv")
End Function
Private Function IDirectory_errorlogFolder() As String
    IDirectory_errorlogFolder = FSO.BuildPath(IDirectory_dataFolder, "errorlog") & AS_FOLDER
End Function
Private Function IDirectory_ErrorlogFile() As String
    IDirectory_ErrorlogFile = FSO.BuildPath(IDirectory_errorlogFolder, "error.log")
End Function
Private Function IDirectory_leaveFolder() As String
    IDirectory_leaveFolder = FSO.BuildPath(IDirectory_dataFolder, "leave") & AS_FOLDER
End Function
Private Function IDirectory_LeaveFile() As String
    IDirectory_LeaveFile = FSO.BuildPath(IDirectory_leaveFolder, "leave.csv")
End Function
Private Function IDirectory_regularFolder() As String
    IDirectory_regularFolder = FSO.BuildPath(IDirectory_dataFolder, "regular") & AS_FOLDER
End Function
Private Function IDirectory_RegularDaysFile() As String
    IDirectory_RegularDaysFile = FSO.BuildPath(IDirectory_regularFolder, "RegularDays.csv")
End Function
Private Function IDirectory_RegularWeekDaysFile() As String
    IDirectory_RegularWeekDaysFile = FSO.BuildPath(IDirectory_regularFolder, "RegularWeekDays.csv")
End Function
Private Function IDirectory_schedulesFolder() As String
    IDirectory_schedulesFolder = FSO.BuildPath(IDirectory_dataFolder, "schedules") & AS_FOLDER
End Function
Private Function IDirectory_schedulesFolderByYear(ByVal WesternYear As String) As String
    IDirectory_schedulesFolderByYear = FSO.BuildPath(IDirectory_schedulesFolder, WesternYear) & AS_FOLDER
End Function
Private Function IDirectory_ScheduleFileByYear(ByVal WesternYear As String) As String
    IDirectory_ScheduleFileByYear = FSO.BuildPath(IDirectory_schedulesFolderByYear(WesternYear), "schedule" & WesternYear & ".csv")
End Function
Private Function IDirectory_writelogFolder() As String
    IDirectory_writelogFolder = FSO.BuildPath(IDirectory_dataFolder, "writelog") & AS_FOLDER
End Function
Private Function IDirectory_headersFolder() As String
    IDirectory_headersFolder = FSO.BuildPath(IDirectory_rootFolder, "headers") & AS_FOLDER
End Function
Private Function IDirectory_ClassHourHeaderFile() As String
    IDirectory_ClassHourHeaderFile = FSO.BuildPath(IDirectory_headersFolder, "ClassHourHeader.csv")
End Function
Private Function IDirectory_EnrollmentHeaderFile() As String
    IDirectory_EnrollmentHeaderFile = FSO.BuildPath(IDirectory_headersFolder, "EnrollmentHeader.csv")
End Function
Private Function IDirectory_ScheduleHeaderFile() As String
    IDirectory_ScheduleHeaderFile = FSO.BuildPath(IDirectory_headersFolder, "ScheduleHeader.csv")
End Function
Private Function IDirectory_holidayFolder() As String
    IDirectory_holidayFolder = FSO.BuildPath(IDirectory_rootFolder, "holiday") & AS_FOLDER
End Function
Private Function IDirectory_HolidayFile() As String
    IDirectory_HolidayFile = FSO.BuildPath(IDirectory_holidayFolder, "Holiday.csv")
End Function
Private Function IDirectory_icoFolder() As String
    IDirectory_icoFolder = FSO.BuildPath(IDirectory_topFolder, "ico") & AS_FOLDER
End Function
Private Function IDirectory_CursorFile() As String
    IDirectory_CursorFile = FSO.BuildPath(IDirectory_icoFolder, "cursor.csv")
End Function
Private Function IDirectory_srcFolder() As String
    IDirectory_srcFolder = FSO.BuildPath(IDirectory_topFolder, "src") & AS_FOLDER
End Function
Private Function IDirectory_templatesFolder() As String
    IDirectory_templatesFolder = FSO.BuildPath(IDirectory_topFolder, "templates") & AS_FOLDER
End Function
Private Function IDirectory_DailyTemplateFile() As String
    IDirectory_DailyTemplateFile = FSO.BuildPath(IDirectory_templatesFolder, "Daily.xlsx")
End Function
Private Function IDirectory_DailyTemplateSettingFile() As String
    IDirectory_DailyTemplateSettingFile = FSO.BuildPath(IDirectory_templatesFolder, "Daily.csv")
End Function
Private Function IDirectory_DailyManualFile() As String
    IDirectory_DailyManualFile = FSO.BuildPath(IDirectory_templatesFolder, "DailyManual.xlsx")
End Function
Private Function IDirectory_MonthlyTemplateFile() As String
    IDirectory_MonthlyTemplateFile = FSO.BuildPath(IDirectory_templatesFolder, "Monthly.xlsx")
End Function
Private Function IDirectory_MonthlyTemplateSettingFile() As String
    IDirectory_MonthlyTemplateSettingFile = FSO.BuildPath(IDirectory_templatesFolder, "Monthly.csv")
End Function
Private Function IDirectory_WeeklyTemplateFile() As String
    IDirectory_WeeklyTemplateFile = FSO.BuildPath(IDirectory_templatesFolder, "Weekly.xlsx")
End Function
Private Function IDirectory_WeeklyTemplateSettingFile() As String
    IDirectory_WeeklyTemplateSettingFile = FSO.BuildPath(IDirectory_templatesFolder, "Weekly.csv")
End Function
Private Function IDirectory_YearlyTemplateFile() As String
    IDirectory_YearlyTemplateFile = FSO.BuildPath(IDirectory_templatesFolder, "Yearly.xlsx")
End Function
Private Function IDirectory_YearlyTemplateSettingFile() As String
    IDirectory_YearlyTemplateSettingFile = FSO.BuildPath(IDirectory_templatesFolder, "Yearly.csv")
End Function
