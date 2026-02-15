VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentConfig"
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

Public Property Get MainStreamMaxGrade() As Long
    MainStreamMaxGrade = CLng(FindValue("MainStreamMaxGrade"))
End Property

Public Property Get MainStreamMaxClass() As Long
    MainStreamMaxClass = CLng(FindValue("MainStreamMaxClass"))
End Property

Public Property Get SpecialMaxGrade() As Long
    SpecialMaxGrade = CLng(FindValue("SpecialMaxGrade"))
End Property

Public Property Get SpecialMaxClass() As Long
    SpecialMaxClass = CLng(FindValue("SpecialMaxClass"))
End Property

Public Property Get DateColumnName() As String
    DateColumnName = CStr(FindValue("TargetDate"))
End Property

Public Property Get MainStreamMaleColumnName(ByVal Grade As Long, ByVal ClassNo As Long) As String
    Dim KeyName As String
    KeyName = BuildKeyName("MainStream", "Male", Grade, ClassNo)
    MainStreamMaleColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get MainStreamFemaleColumnName(ByVal Grade As Long, ByVal ClassNo As Long) As String
    Dim KeyName As String
    KeyName = BuildKeyName("MainStream", "Female", Grade, ClassNo)
    MainStreamFemaleColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get SpecialMaleColumnName(ByVal Grade As Long, ByVal ClassNo As Long) As String
    Dim KeyName As String
    KeyName = BuildKeyName("Special", "Male", Grade, ClassNo)
    SpecialMaleColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get SpecialFemaleColumnName(ByVal Grade As Long, ByVal ClassNo As Long) As String
    Dim KeyName As String
    KeyName = BuildKeyName("Special", "Female", Grade, ClassNo)
    SpecialFemaleColumnName = CStr(FindValue(KeyName))
End Property

Public Property Get TransferColumnName() As String
    TransferColumnName = CStr(FindValue("Transfer"))
End Property

Public Property Get RemarksColumnName() As String
    RemarksColumnName = CStr(FindValue("Remarks"))
End Property

Private Function BuildKeyName(ByVal ClassType As String, ByVal Gender As String, ByVal Grade As Long, ByVal ClassNo As Long) As String
    BuildKeyName = ClassType & "_" & CStr(Grade) & "_" & CStr(ClassNo) & "_" & Gender
End Function

Private Function FindValue(ByVal KeyName As String) As Variant
    Dim KeyValue As Dom_KeyValue
    For Each KeyValue In This.Keys
        If KeyValue.Name = KeyName Then
            FindValue = KeyValue.Value
            Exit Function
        End If
    Next
    Err.Raise NotFoundConfigKey, "App_EnrollmentConfig", "Config key not found: " & KeyName
End Function

