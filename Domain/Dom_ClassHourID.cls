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
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    This.ClassHourType = ClassHourType
    Set This.Subject = Subject
    Set This.Period = Period
    Set This.Grade = Grade
End Sub

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    Dim Result As Dom_ClassHourID
    Set Result = New Dom_ClassHourID
    Result.InitializeInternal ClassHourType, Subject, Period, Grade
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_ClassHourID) As Boolean
    If Other Is Nothing Then Exit Function
    If This.ClassHourType <> Other.ClassHourType Then Exit Function
    If Not This.Subject.Equals(Other.Subject) Then Exit Function
    If Not This.Period.Equals(Other.Period) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    Equals = True
End Function

Public Function ToKey() As String
    Dim ClassHourTypeKey As String
    Select Case This.ClassHourType
    Case Dom_ClassHourTypeEnum.Plan: ClassHourTypeKey = jpnKeikaku
    Case Dom_ClassHourTypeEnum.Execution: ClassHourTypeKey = jpnJisshi
    End Select
    If This.Subject Is Nothing Then
        ToKey = ClassHourTypeKey & charUnderScore & This.Period.ToKey & charUnderScore & This.Grade.ToKey
    ElseIf This.Period Is Nothing Then
        ToKey = ClassHourTypeKey & charUnderScore & This.Subject.ToKey & charUnderScore & This.Grade.ToKey
    End If
End Function

Public Property Get ClassHourType() As Dom_ClassHourType
    ClassHourType = This.ClassHourType
End Property

Public Property Get Subject() As Dom_Subject
    Set Subject = This.Subject
End Property

Public Property Get Period() As Dom_Period
    Set Period = This.Period
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property
