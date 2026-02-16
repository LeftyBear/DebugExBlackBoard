VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.VO"
Option Explicit

Private Type Member
    Keys As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal KeyValues As VBA.Collection)
    Set This.Keys = KeyValues
End Sub

Public Property Get MaxGrade() As Long
    MaxGrade = CLng(FindValue("MainStreamMaxGrade"))
End Property

Public Property Get MaxSubject() As Long
    MaxSubject = CLng(FindValue("MaxSubject"))
End Property

Public Property Get MaxPeriod() As Long
    MaxPeriod = CLng(FindValue("MaxPeriod"))
End Property

Public Property Get TargetGradeOf(ByVal SubjectName As String, ByVal Index As Long) As Long
    Dim KeyName As String
    KeyName = BuildTargetGradeKey(SubjectName, Index)
    TargetGradeOf = CLng(FindValue(KeyName))
End Property

Public Property Get DateColumnName() As String
    DateColumnName = CStr(FindValue("TargetDate"))
End Property

Public Property Get ExeClassHourColumnName(ByVal Grade As Long, ByVal SubjectName As String, Optional ByVal Index As Long) As String
    Dim KeyName As String
    KeyName = BuildClassHourKey("Exe", Grade, SubjectName, Index)
    ExeClassHourColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get PlanClassHourColumnName(ByVal Grade As Long, ByVal SubjectName As String, Optional ByVal Index As Long) As String
    Dim KeyName As String
    KeyName = BuildClassHourKey("Plan", Grade, SubjectName, Index)
    PlanClassHourColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get ExeTimeTableColumnName(ByVal Grade As Long, ByVal PeriodName As String, Optional ByVal Index As Long) As String
    Dim KeyName As String
    KeyName = BuildTimeTableKey("Exe", Grade, PeriodName, Index)
    ExeTimeTableColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get PlanTimeTableColumnName(ByVal Grade As Long, ByVal PeriodName As String, Optional ByVal Index As Long) As String
    Dim KeyName As String
    KeyName = BuildTimeTableKey("Plan", Grade, PeriodName, Index)
    PlanTimeTableColumnName = CStr(FindValue(KeyName))
End Property

Private Function BuildTargetGradeKey(ByVal SubjectName As String, ByVal Index As Long) As String
    BuildTargetGradeKey = SubjectName & CStr(Index) & "_TargetGrade"
End Function

Private Function BuildClassHourKey(ByVal Prefix As String, ByVal Grade As Long, ByVal SubjectName As String, Optional ByVal Index As Long) As String
    If Index = 0 Then
        BuildClassHourKey = Prefix & "_" & SubjectName & "_" & CStr(Grade)
    Else
        BuildClassHourKey = Prefix & "_" & SubjectName & CStr(Index) & "_" & CStr(Grade)
    End If
End Function

Private Function BuildTimeTableKey(ByVal Prefix As String, ByVal Grade As Long, ByVal PeriodName As String, ByVal Index As Long) As String
    If Index = 0 Then
        BuildTimeTableKey = Prefix & "_" & PeriodName & "_" & CStr(Grade)
    Else
        BuildTimeTableKey = Prefix & "_" & PeriodName & CStr(Index) & "_" & CStr(Grade)
    End If
End Function

Private Function FindValue(ByVal KeyName As String) As Variant
    Dim KeyValue As App_KeyValue
    For Each KeyValue In This.Keys
        If KeyValue.Name = KeyName Then
            FindValue = KeyValue.Value
            Exit Function
        End If
    Next
    Err.Raise NotFoundConfigKey, "App_EnrollmentConfig", "Config key not found: " & KeyName
End Function

