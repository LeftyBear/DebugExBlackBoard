VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EntityFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Public Function Resolve(ByVal EntityName As Dom_EntityTypeEnum, ByVal SchoolYear As Dom_SchoolYear) As String
    If EntityName = Dom_EntityTypeEnum.ClassHour Then
        Resolve = "ClassHour" & SchoolYear.ToKey & ".csv"
    ElseIf EntityName = Dom_EntityTypeEnum.Enrollment Then
        Resolve = "Enrollment" & SchoolYear.ToKey & ".csv"
    ElseIf EntityName = Dom_EntityTypeEnum.Schedule Then
        Resolve = "Schedule" & SchoolYear.ToKey & ".csv"
    End If
End Function
