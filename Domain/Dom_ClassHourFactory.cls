VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    ValueObjectFactory  As Dom_ValueObjectFactory
    Structure           As Dom_SchoolStructure
End Type

Private This As Member

Public Sub Initialize(ByVal ValueObjectFactory As Dom_ValueObjectFactory)
    Set This.ValueObjectFactory = ValueObjectFactory
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Create(ByVal Row As App_ClassHourRow) As Dom_ClassHour
    Dim DateID As Dom_ClassHourID
    Set DateID = This.ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim ID As Dom_ClassHourID
    Set ID = CreateID(Row)
    Dim Value As Dom_ClassHourValue
    Set Value = This.ValueObjectFactory.CreateClassHourValue(Row.Value)
    If ID Is Nothing Then Exit Function
    Dim Entity As Dom_ClassHour
    Set Entity = New Dom_ClassHour
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function

Private Function CreateID(ByVal Row As App_ClassHourRow) As Dom_ClassHourID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.ValueObjectFactory.CreateClassHourType(Row.TypeCode)
    If ClassHourType.Value = 0 Then Exit Function
    Dim Subject As Dom_Subject
    Set Subject = This.Structure.SubjectByName(Row.Subject)
    If Subject.Index = 0 Then Exit Function
    Dim Grade As Dom_Grade
    Set Grade = This.ValueObjectFactory.CreateGrade(Row.Grade)
    If Grade.Value = 0 Then Exit Function
    Set CreateID = This.ValueObjectFactory.CreateClassHourID(ClassHourType, Subject, Grade)
End Function
