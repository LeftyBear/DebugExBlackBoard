VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_IClassHourQueryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.QueryServiceInterface"
Option Explicit

Public Function GetByDate(ByVal TargetDate As Date) As VBA.Collection
End Function

Public Function GetSubjectTotal(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long, ByVal Subject As String) As Long
End Function

Public Function GetHoursCount(ByVal TargetDate As Date, ByVal TypeCode As Dom_StreamType, ByVal Grade As Long) As Long
End Function
