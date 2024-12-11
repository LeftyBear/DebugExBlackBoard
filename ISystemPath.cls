VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISystemPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public topFolder As String
Public Function rootFolder() As String
End Function
Public Function dataFolder() As String
End Function
Public Function backupFolder() As String
End Function
Public Function cacheFolder() As String
End Function
Public Function AdministerFile() As String
End Function
Public Function OnTimeFile() As String
End Function
Public Function PasswordFile() As String
End Function
Public Function PasswordSkipFile() As String
End Function
Public Function classhoursFolder() As String
End Function
Public Function classhoursFolderByYear(ByVal WesternYear As String) As String
End Function
Public Function ClasshoursFileByYear(ByVal WesternYear As String) As String
End Function
Public Function daydutyFolder() As String
End Function
Public Function DayDutyFile() As String
End Function
Public Function enrollmentsFolder() As String
End Function
Public Function enrollmentsFolderByYear(ByVal WesternYear As String) As String
End Function
Public Function EnrollmentsFileByYear(ByVal WesternYear As String) As String
End Function
Public Function errorlogFolder() As String
End Function
Public Function leaveFolder() As String
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
Public Function SchedulesFileByYear(ByVal WesternYear As String) As String
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
Public Function DailyFile() As String
End Function
Public Function DailyManualFile() As String
End Function
Public Function MonthlyFile() As String
End Function
Public Function WeeklyFile() As String
End Function
Public Function YearlyFile() As String
End Function
