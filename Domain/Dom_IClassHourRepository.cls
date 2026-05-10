VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_IClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Add(ByVal Item As Dom_DailyClassHour)
End Sub

Public Function FindByDate(ByVal TargetDate As Date) As Dom_DailyClassHour
End Function

Public Function GetAll() As VBA.Collection
End Function

Public Function Replace(ByVal Aggregate As Dom_DailyClassHour) As VBA.Collection
End Function

Public Function ReadAll(ByVal SchoolYear As Long) As VBA.Collection
End Function
