VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    FixDate         As Date
    Values()        As Long
    TimeTables()    As String
End Type

Private This As Member

Public Sub Initialize(ByVal FixDate As Date, ByRef Values() As Long, ByRef TimeTables() As String)
    This.FixDate = FixDate
    This.Values = Values
    This.TimeTables = TimeTables
End Sub

Public Property Get FixDate() As Date
    FixDate = This.FixDate
End Property

Public Property Get Value(ByVal ClassHourType As Long, ByVal SubjectIndex As Long, ByVal Grade As Long) As Long
    Value = This.Values(ClassHourType, SubjectIndex, Grade)
End Property

Public Property Get TimeTable(ByVal ClassHourType As Long, ByVal Period As Long, ByVal Grade As Long) As String
    TimeTable = This.TimeTables(ClassHourType, Period, Grade)
End Property
