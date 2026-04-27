VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Structure       As Dom_SchoolStructure
    EntityFactory   As App_EntityFactory
    VOFactory       As App_ValueObjectFactory
End Type

Private This As Member

Private Sub Initialize(ByVal EntityFactory As App_EntityFactory, ByVal VOFactory As App_ValueObjectFactory)
    Set This.EntityFactory = EntityFactory
    Set This.VOFactory = VOFactory
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Create(ByVal NormDate As Date, ByVal Row As App_ClassHourRow) As Dom_ClassHour
    Dim ID As Dom_ClassHourID
    Set ID = CreateID(Row)
    If ID.IsNoMeaning Then Exit Function
    Dim DateID As Dom_ClassHourID
    Set DateID = This.VOFactory.CreateDateID(NormDate)
    Dim Value As Dom_ClassHourValue
    Set Value = This.VOFactory.CreateClassHourValue(Row.Value)
    Set Create = This.EntityFactory.CreateClassHour(DateID, ID, Value)
End Function

Private Function CreateID(ByVal Row As App_ClassHourRow) As Dom_ClassHourID
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.VOFactory.CreateClassHourType(Row.TypeCode)
    Dim Subject As Dom_Subject
    Set Subject = This.Structure.ResolveSubject(Row.Subject)
    Dim Grade As Dom_Grade
    Set Grade = This.VOFactory.CreateGrade(Row.Grade)
    Set CreateID = This.VOFactory.CreateClassHourID(ClassHourType, Subject, Grade)
End Function
