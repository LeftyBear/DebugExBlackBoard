VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EntityFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Resolver")
Option Explicit

Public Function Resolve(ByVal EntityName As Dom_EntityTypeEnum, ByVal SchoolYear As Dom_SchoolYear) As String
    Select Case EntityName
    Case Dom_EntityTypeEnum.ClassHour
        Resolve = "ClassHour" & SchoolYear.ToKey & ".csv"
    Case Dom_EntityTypeEnum.Enrollment
        Resolve = "Enrollment" & SchoolYear.ToKey & ".csv"
    Case Dom_EntityTypeEnum.Schedule
        Resolve = "Schedule" & SchoolYear.ToKey & ".csv"
    End Select
End Function
