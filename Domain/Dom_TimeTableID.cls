VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    ClassHourType   As Dom_ClassHourType
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Friend Sub InitializeInternal(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    Set This.ClassHourType = ClassHourType
    Set This.Period = Period
    Set This.Grade = Grade
End Sub

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableID
    Dim Result As Dom_TimeTableID
    Set Result = New Dom_TimeTableID
    Result.InitializeInternal ClassHourType, Period, Grade
    Set Create = Result
End Function

Public Function IsNoMeaning() As Boolean
    If 0 < This.Period.Value And 0 < This.Grade.Value Then Exit Function
    IsNoMeaning = True
End Function

Public Function Equals(ByVal Other As Dom_TimeTableID) As Boolean
    If Other Is Nothing Then Exit Function
    If Not This.ClassHourType.Equals(Other.ClassHourType) Then Exit Function
    If Not This.Period.Equals(Other.Period) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    Equals = True
End Function

Public Function ToKey() As String
    ToKey = This.ClassHourType.ToKey & charUnderScore & This.Period.ToKey & charUnderScore & This.Grade.ToKey
End Function

Public Property Get ClassHourType() As Dom_ClassHourType
    Set ClassHourType = This.ClassHourType
End Property

Public Property Get Period() As Dom_Period
    Set Period = This.Period
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property
