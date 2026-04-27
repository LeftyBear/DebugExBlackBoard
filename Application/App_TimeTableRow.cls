VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTO")
Option Explicit

Private Type Member
    TypeCode    As Long
    Period      As Long
    Grade       As Long
    Value       As String
End Type

Private This As Member

Public Sub SetTypeCode(ByVal TypeCode As Long)
    This.TypeCode = TypeCode
End Sub

Public Sub SetPeriod(ByVal Period As Long)
    This.Period = Period
End Sub

Public Sub SetGrade(ByVal Grade As Long)
    This.Grade = Grade
End Sub

Public Sub SetValue(ByVal Value As String)
    This.Value = Value
End Sub

Public Property Get TypeCode() As Long
    TypeCode = This.TypeCode
End Property

Public Property Get Period() As Long
    Period = This.Period
End Property

Public Property Get Grade() As Long
    Grade = This.Grade
End Property

Public Property Get Value() As String
    Value = This.Value
End Property
