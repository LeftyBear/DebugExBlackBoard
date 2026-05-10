VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_IEnrollmentQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.QueryServiceInterface"
Option Explicit

Public Function GetByDate(ByVal TargetDate As Date) As VBA.Collection
End Function

Public Function GetMaleTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long) As Long
End Function

Public Function GetFemaleTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long) As Long
End Function

Public Function GetClassTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long, ByVal ClassNo As Long) As Long
End Function
