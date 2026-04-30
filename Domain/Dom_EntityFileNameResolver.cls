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

Public Function Resolve(ByVal TypeCode As Dom_EntityTypePolicy, ByVal SchoolYear As Long) As String
    If TypeCode = Dom_EntityTypePolicy.ClassHour Then
        Resolve = "ClassHour" & SchoolYear & ".csv"
    ElseIf TypeCode = Dom_EntityTypePolicy.Enrollment Then
        Resolve = "Enrollment" & SchoolYear & ".csv"
    ElseIf TypeCode = Dom_EntityTypePolicy.Schedule Then
        Resolve = "Schedule" & SchoolYear & ".csv"
    End If
End Function
