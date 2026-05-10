VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_IEnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Save(ByVal SchoolYear As Long)
End Sub

Public Sub Replace(ByVal TargetDate As Date, ByVal NewAggregate As Dom_DailyEnrollment)
End Sub

Public Function FindByDate(ByVal TargetDate As Date) As Dom_DailyEnrollment
End Function
