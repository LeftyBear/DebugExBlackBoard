VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EntityFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Resolver")
Option Explicit

Public Function Resolve(ByVal EntityName As Dom_Entity, ByVal SchoolYear As Long) As String
    Select Case EntityName
    Case Dom_Entity.Enrollment
        Resolve = "Enrollment" & CStr(SchoolYear) & ".csv"
    Case Dom_Entity.ClassHour
        Resolve = "ClassHour" & CStr(SchoolYear) & ".csv"
    End Select
End Function
