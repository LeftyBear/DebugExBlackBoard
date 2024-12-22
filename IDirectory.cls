VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, VariableNotAssigned, MoveFieldCloserToUsage, EncapsulatePublicField
'@Folder("Directory")
Option Explicit
Public TopFolderPath As String
Public Function topFolder() As String
End Function
Public Function rootFolder() As String
End Function
Public Function dataFolder() As String
End Function
Public Function backupFolder() As String
End Function
Public Function cacheFolder() As String
End Function
Public Function AdministorFile() As String
End Function
Public Function SwappingCheckFile() As String
End Function
Public Function PasswordFile() As String
End Function
Public Function PasswordSkipFile() As String
End Function
Public Function classhoursFolder() As String
End Function
Public Function classhoursFolderByYear(ByVal WesternYear As String) As String
End Function
Public Function ClassHourFileByYear(ByVal WesternYear As String) As String
End Function
Public Function daydutyFolder() As String
End Function
Public Function DayDutyFile() As String
End Function
Public Function enrollmentsFolder() As String
End Function
Public Function enrollmentsFolderByYear(ByVal WesternYear As String) As String
End Function
Public Function EnrollmentFileByYear(ByVal WesternYear As String) As String
End Function
Public Function errorlogFolder() As String
End Function
Public Function ErrorlogFile() As String
End Function
Public Function leaveFolder() As String
End Function
Public Function LeaveFile() As String
End Function
Public Function regularFolder() As String
End Function
Public Function RegularDaysFile() As String
End Function
Public Function RegularWeekDaysFile() As String
End Function
Public Function schedulesFolder() As String
End Function
Public Function schedulesFolderByYear(ByVal WesternYear As String) As String
End Function
Public Function ScheduleFileByYear(ByVal WesternYear As String) As String
End Function
Public Function writelogFolder() As String
End Function
Public Function headersFolder() As String
End Function
Public Function ClassHourHeaderFile() As String
End Function
Public Function EnrollmentHeaderFile() As String
End Function
Public Function ScheduleHeaderFile() As String
End Function
Public Function holidayFolder() As String
End Function
Public Function HolidayFile() As String
End Function
Public Function icoFolder() As String
End Function
Public Function CursorFile() As String
End Function
Public Function srcFolder() As String
End Function
Public Function templatesFolder() As String
End Function
Public Function DailyTemplateFile() As String
End Function
Public Function DailyTemplateSettingFile() As String
End Function
Public Function DailyManualFile() As String
End Function
Public Function MonthlyTemplateFile() As String
End Function
Public Function MonthlyTemplateSettingFile() As String
End Function
Public Function WeeklyTemplateFile() As String
End Function
Public Function WeeklyTemplateSettingFile() As String
End Function
Public Function YearlyTemplateFile() As String
End Function
Public Function YearlyTemplateSettingFile() As String
End Function

