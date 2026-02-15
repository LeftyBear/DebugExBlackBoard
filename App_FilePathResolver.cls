VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_FilePathResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Application.Resolver"
Option Explicit

Private Type Member
    BaseFolderPath As String 'FSO.GetParentFolderName(FSO.GetParentFolderName(ThisWorkbook.Path))
End Type

Private This As Member
Private FSO As New Scripting.FileSystemObject

Public Sub Initialize(ByVal BaseFolderPath As String)
    This.BaseFolderPath = BaseFolderPath
End Sub

'TopFolder--------------------------------------------------------------------------------------------
Public Function RootFolder() As String
    RootFolder = FSO.BuildPath(This.BaseFolderPath, "root") & Util_Character.BackSlash
End Function

'   >RootFolder---------------------------------------------------------------------------------------
Public Function DataFolder() As String
    DataFolder = FSO.BuildPath(RootFolder, "data") & Util_Character.BackSlash
End Function

'       >ConfigFolder-----------------------------------------------------------------------------------
Public Function ConfigFolder() As String
    ConfigFolder = FSO.BuildPath(RootFolder, "config") & Util_Character.BackSlash
End Function

'       >---------------------------------------------------------------------------------------------
Public Function DesignFile() As String
    DesignFile = FSO.BuildPath(ConfigFolder, "design.ini")
End Function

'       >DataFolder-----------------------------------------------------------------------------------
Public Function BackupFolder() As String
    BackupFolder = FSO.BuildPath(DataFolder, "backup") & Util_Character.BackSlash
End Function

'       >---------------------------------------------------------------------------------------------
Public Function CacheFolder() As String
    CacheFolder = FSO.BuildPath(DataFolder, "cache") & Util_Character.BackSlash
End Function

'           >CacheFolder------------------------------------------------------------------------------
Public Function AdministratorFile() As String
    AdministratorFile = FSO.BuildPath(CacheFolder, "Administrator.txt")
End Function

Public Function ApplyingCheckFile() As String
    ApplyingCheckFile = FSO.BuildPath(CacheFolder, "OnTime.log")
End Function

Public Function PasswordFile() As String
    PasswordFile = FSO.BuildPath(CacheFolder, "Password.txt")
End Function

Public Function PasswordSkipsFile() As String
    PasswordSkipsFile = FSO.BuildPath(CacheFolder, "PasswordSkips.txt")
End Function

'       >----------------------------------------------------------------------------------------------
Public Function ErrorlogFolder() As String
    ErrorlogFolder = FSO.BuildPath(DataFolder, "errorlog") & Util_Character.BackSlash
End Function

'           >ErrorlogFolder------------------------------------------------------------------------------
Public Function ErrorlogFile() As String
    ErrorlogFile = FSO.BuildPath(ErrorlogFolder, "Error.log")
End Function

'       >----------------------------------------------------------------------------------------------
Public Function SchoolYearsFolder() As String
    SchoolYearsFolder = FSO.BuildPath(DataFolder, "schoolyears") & Util_Character.BackSlash
End Function

'           >schoolyearsFolder-----------------------------------------------------------------------------
Public Function SchoolYearFolderByYear(ByVal TargetDate As Date) As String
    SchoolYearFolderByYear = FSO.BuildPath(SchoolYearsFolder, ResolveSchoolYear(TargetDate)) & Util_Character.BackSlash
End Function

Public Function ClassHourFileByYear(ByVal TargetDate As Date) As String
    ClassHourFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "ClassHour" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function OnDutyFileByYear(ByVal TargetDate As Date) As String
    OnDutyFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "DayDuty" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function EnrollmentFileByYear(ByVal TargetDate As Date) As String
    EnrollmentFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "Enrollment" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function KidsGroupActivitiesHoursFileByYear(ByVal TargetDate As Date) As String
    KidsGroupActivitiesHoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "KidsGroupActivitiesHours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function KidsGroupActivitiesScheduleFileByYear(ByVal TargetDate As Date) As String
    KidsGroupActivitiesScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "KidsGroupActivitiesSchedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function LeaveFileByYear(ByVal TargetDate As Date) As String
    LeaveFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "Leave" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function RegularDayFileByYear(ByVal TargetDate As Date) As String
    RegularDayFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "RegularDay" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function RegularWeekDayFileByYear(ByVal TargetDate As Date) As String
    RegularWeekDayFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "RegularWeekDay" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function ScheduleFileByYear(ByVal TargetDate As Date) As String
    ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolHolidayFileByYear(ByVal TargetDate As Date) As String
    SchoolHolidayFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolHoliday" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function StandardClassHourFileByYear(ByVal TargetDate As Date) As String
    StandardClassHourFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "StandardClassHour" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent01HoursFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent01HoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent01Hours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent01ScheduleFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent01ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent01Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent02HoursFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent02HoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent02Hours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent02ScheduleFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent02ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent02Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent03HoursFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent03HoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent03Hours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent03ScheduleFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent03ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent03Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent04HoursFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent04HoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent04Hours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent04ScheduleFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent04ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent04Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent05HoursFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent05HoursFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent05Hours" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function SchoolEvent05ScheduleFileByYear(ByVal TargetDate As Date) As String
    SchoolEvent05ScheduleFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "SchoolEvent05Schedule" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function WeeklyClassHourFileByYear(ByVal TargetDate As Date) As String
    WeeklyClassHourFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "WeeklyClassHour" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function WeeklyTimeTableFileByYear(ByVal TargetDate As Date) As String
    WeeklyTimeTableFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "WeeklyTimeTable" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

Public Function YearlyClassHourFileByYear(ByVal TargetDate As Date) As String
    YearlyClassHourFileByYear = FSO.BuildPath(SchoolYearFolderByYear(TargetDate), "YearlyClassHour" & ResolveSchoolYear(TargetDate) & ".csv")
End Function

'       >----------------------------------------------------------------------------------------------
Public Function WritelogFolder() As String
    WritelogFolder = FSO.BuildPath(DataFolder, "writelog") & Util_Character.BackSlash
End Function

'   >--------------------------------------------------------------------------------------------------
Public Function HolidayFolder() As String
    HolidayFolder = FSO.BuildPath(RootFolder, "holiday") & Util_Character.BackSlash
End Function

'       >HolidayFolder--------------------------------------------------------------------------------
Public Function HolidayFile() As String
    HolidayFile = FSO.BuildPath(HolidayFolder, "syukujitsu.csv")
End Function

'   >--------------------------------------------------------------------------------------------------
Public Function HeadersFolder() As String
    HeadersFolder = FSO.BuildPath(RootFolder, "headers") & Util_Character.BackSlash
End Function

'       >HeadersFolder--------------------------------------------------------------------------------
Public Function ClassHourHeaderFile() As String
    ClassHourHeaderFile = FSO.BuildPath(HeadersFolder, "ClassHourHeader.csv")
End Function

Public Function OnDutyHeaderFile() As String
    OnDutyHeaderFile = FSO.BuildPath(HeadersFolder, "OnDutyHeader.csv")
End Function

Public Function EnrollmentHeaderFile() As String
    EnrollmentHeaderFile = FSO.BuildPath(HeadersFolder, "EnrollmentHeader.csv")
End Function

Public Function EnrollmentHeaderMappingINI() As String
    EnrollmentHeaderMappingINI = FSO.BuildPath(HeadersFolder, "EnrollmentHeader.ini")
End Function

Public Function LeaveHeaderFile() As String
    LeaveHeaderFile = FSO.BuildPath(HeadersFolder, "LeaveHeader.csv")
End Function

Public Function RegularDayHeaderFile() As String
    RegularDayHeaderFile = FSO.BuildPath(HeadersFolder, "RegularDayHeader.csv")
End Function

Public Function RegularWeekDayHeaderFile() As String
    RegularWeekDayHeaderFile = FSO.BuildPath(HeadersFolder, "RegularWeekDayHeader.csv")
End Function

Public Function ScheduleHeaderFile() As String
    ScheduleHeaderFile = FSO.BuildPath(HeadersFolder, "ScheduleHeader.csv")
End Function

Public Function SchoolEventCommonHeaderFile() As String
    SchoolEventCommonHeaderFile = FSO.BuildPath(HeadersFolder, "SchoolEventCommonHeader.csv")
End Function

Public Function SchoolHolidayHeaderFile() As String
    SchoolHolidayHeaderFile = FSO.BuildPath(HeadersFolder, "SchoolHolidayHeader.csv")
End Function

Public Function StandardClassHourHeaderFile() As String
    StandardClassHourHeaderFile = FSO.BuildPath(HeadersFolder, "StandardClassHourHeader.csv")
End Function

Public Function WeeklyClassHourHeaderFile() As String
    WeeklyClassHourHeaderFile = FSO.BuildPath(HeadersFolder, "WeeklyClassHourHeader.csv")
End Function

Public Function WeeklyTimeTableHeaderFile() As String
    WeeklyTimeTableHeaderFile = FSO.BuildPath(HeadersFolder, "WeeklyTimeTableHeader.csv")
End Function

Public Function YearlyClassHourHeaderFile() As String
    YearlyClassHourHeaderFile = FSO.BuildPath(HeadersFolder, "YearlyClassHourHeader.csv")
End Function

'   >--------------------------------------------------------------------------------------------------
Public Function icoFolder() As String
    icoFolder = FSO.BuildPath(RootFolder, "ico") & Util_Character.BackSlash
End Function

'       >IcoFolder------------------------------------------------------------------------------------
Public Function CursorFile() As String
    CursorFile = FSO.BuildPath(icoFolder, "Cursor.ico")
End Function

'   >--------------------------------------------------------------------------------------------------
Public Function srcFolder() As String
    srcFolder = FSO.BuildPath(RootFolder, "src") & Util_Character.BackSlash
End Function

'   >--------------------------------------------------------------------------------------------------
Public Function TemplatesFolder() As String
    TemplatesFolder = FSO.BuildPath(RootFolder, "templates") & Util_Character.BackSlash
End Function

'       >TemplatesFolder------------------------------------------------------------------------------
Public Function DailyTemplateFile() As String
    DailyTemplateFile = FSO.BuildPath(TemplatesFolder, "Daily.xlsx")
End Function

Public Function DailyTemplateSettingFile() As String
    DailyTemplateSettingFile = FSO.BuildPath(TemplatesFolder, "Daily.csv")
End Function

Public Function DailyManualFile() As String
    DailyManualFile = FSO.BuildPath(TemplatesFolder, "DailyManual.xlsx")
End Function

Public Function MonthlyTemplateFile() As String
    MonthlyTemplateFile = FSO.BuildPath(TemplatesFolder, "Monthly.xlsx")
End Function

Public Function MonthlyTemplateSettingFile() As String
    MonthlyTemplateSettingFile = FSO.BuildPath(TemplatesFolder, "Monthly.csv")
End Function

Public Function WeeklyTemplateFile() As String
    WeeklyTemplateFile = FSO.BuildPath(TemplatesFolder, "Weekly.xlsx")
End Function

Public Function WeeklyTemplateSettingFile() As String
    WeeklyTemplateSettingFile = FSO.BuildPath(TemplatesFolder, "Weekly.csv")
End Function

Public Function YearlyTemplateFile() As String
    YearlyTemplateFile = FSO.BuildPath(TemplatesFolder, "Yearly.xlsx")
End Function

Public Function YearlyTemplateSettingFile() As String
    YearlyTemplateSettingFile = FSO.BuildPath(TemplatesFolder, "Yearly.csv")
End Function
