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
    ClassHourType   As Dom_ClassHourType
    Subject         As Dom_Subject
    Grade           As Dom_Grade
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade)
    Set This.ClassHourType = ClassHourType
    Set This.Subject = Subject
    Set This.Grade = Grade
End Sub

Public Property Get ClassHourType() As Dom_ClassHourType
    Set ClassHourType = This.ClassHourType
End Property

Public Property Get Subject() As Dom_Subject
    Set Subject = This.Subject
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Public Function ToKey() As String
    ToKey = This.ClassHourType.ToKey & charUnderScore & This.Subject.ToKey & charUnderScore & This.Grade.ToKey
End Function

Public Function Equals(ByVal Other As Dom_ClassHourID) As Boolean
    If Not This.ClassHourType.Equals(Other.ClassHourType) Then Exit Function
    If Not This.Subject.Equals(Other.Subject) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    Equals = True
End Function

Public Function IsPlan() As Boolean
    IsPlan = (This.ClassHourType.Value = Dom_ClassHourTypeEnum.Plan)
End Function

Public Function IsExecution() As Boolean
    IsExecution = (This.ClassHourType.Value = Dom_ClassHourTypeEnum.Execution)
End Function
