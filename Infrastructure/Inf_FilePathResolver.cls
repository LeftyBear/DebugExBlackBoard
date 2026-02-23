VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_FilePathResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service")
Option Explicit

Public Enum Inf_EntityFileType
    ClassHourFile
    EnrollmentFile
End Enum

Private Type Member
    Provider As Inf_WorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Provider As Inf_WorkbookPathProvider)
    Set This.Provider = Provider
End Sub

Public Function Resolve(ByVal FileType As Inf_EntityFileType, ByVal SchoolYear As Long) As String
    Dim BasePath As String
    BasePath = This.Provider.GetBaseFolderPath
    Dim YearFolder As String
    YearFolder = BuildYearFolder(SchoolYear)
    Dim FileName As String
    FileName = BuildFileName(FileType, SchoolYear)
    Resolve = BasePath & charBackSlash & YearFolder & charBackSlash & FileName
End Function

Private Function BuildYearFolder(ByVal SchoolYear As Long) As String
    BuildYearFolder = CStr(SchoolYear)
End Function

Private Function BuildFileName(ByVal FileType As Inf_EntityFileType, ByVal SchoolYear As Long) As String
    Select Case FileType
    Case EnrollmentFile
        BuildFileName = "Enrollment" & CStr(SchoolYear) & ".csv"
    Case ClassHourFile
        BuildFileName = "ClassHour" & CStr(SchoolYear) & ".csv"
    End Select
End Function
