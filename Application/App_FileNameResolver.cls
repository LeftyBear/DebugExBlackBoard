VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_FileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.Resolver
Option Explicit

Public Enum App_FileEntityType
    Enrollment
    ClassHour
End Enum

Public Function Resolve(ByVal SchoolYearKey As App_SchoolYearKey, ByVal EntityType As App_FileEntityType) As String
    Select Case EntityType
    Case Enrollment
        Resolve = "Enrollment" & SchoolYearKey.Value & ".csv"
    Case ClassHour
        Resolve = "ClassHour" & SchoolYearKey.Value & ".csv"
    End Select
End Function
