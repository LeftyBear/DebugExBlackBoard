VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    ClassHourType   As String
    SubjectName     As String
    Grade           As String
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourType As String, ByVal SubjectName As String, ByVal Grade As String)
    This.ClassHourType = ClassHourType
    This.SubjectName = SubjectName
    This.Grade = Grade
End Sub

Public Property Get ClassHourType() As String
    ClassHourType = This.ClassHourType
End Property

Public Property Get SubjectName() As String
    SubjectName = This.SubjectName
End Property

Public Property Get Grade() As String
    Grade = This.Grade
End Property
