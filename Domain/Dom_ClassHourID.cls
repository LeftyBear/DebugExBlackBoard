VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    TypeCode    As Dom_ClassHourType
    Subject     As Dom_Subject
    Grade       As Dom_Grade
End Type

Private This As Member

Public Sub Initialize(ByVal TypeCode As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade)
    Set This.TypeCode = TypeCode
    Set This.Subject = Subject
    Set This.Grade = Grade
End Sub

Public Property Get TypeCode() As Dom_ClassHourType
    Set TypeCode = This.TypeCode
End Property

Public Property Get Subject() As Dom_Subject
    Set Subject = This.Subject
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Public Function ToKey() As String
    Dim TypeCodeMapper As Inf_ClassHourTypeMapper
    Set TypeCodeMapper = New Inf_ClassHourTypeMapper
    Dim GradeMapper As Inf_GradeMapper
    Set GradeMapper = New Inf_GradeMapper
    ToKey = TypeCodeMapper.ToRaw(This.TypeCode) & charUnderScore & This.Subject.Name & charUnderScore & GradeMapper.ToRaw(This.Grade)
End Function

Public Function Equals(ByVal Other As Dom_ClassHourID) As Boolean
    If Not This.TypeCode.Equals(Other.TypeCode) Then Exit Function
    If Not This.Subject.Equals(Other.Subject) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    Equals = True
End Function

Public Function IsPlan() As Boolean
    IsPlan = (This.TypeCode.Value = Dom_ClassHourTypePolicy.Plan)
End Function

Public Function IsExecution() As Boolean
    IsExecution = (This.TypeCode.Value = Dom_ClassHourTypePolicy.Execution)
End Function
