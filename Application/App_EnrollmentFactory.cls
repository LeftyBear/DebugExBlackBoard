VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Factory As App_ValueObjectFactory
End Type

Private This As Member

Private Sub Initialize(ByVal Factory As App_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByVal NormDate As Date, ByVal Row As App_EnrollmentRow) As Dom_Enrollment
    Dim ID As Dom_EnrollmentID
    Set ID = CreateID(Row)
    If ID.IsNoMeaning Then Exit Function
    Dim DateID As Dom_EnrollmentID
    Set DateID = This.Factory.CreateDateID(NormDate)
    Dim Value As Dom_EnrollmentValue
    Set Value = This.Factory.CreateEnrollmentValue(Row.Value)
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function

Private Function CreateID(ByVal Row As App_EnrollmentRow) As Dom_EnrollmentID
    Dim StreamType As Dom_StreamType
    Set StreamType = This.Factory.CreateStreamType(Row.TypeCode)
    Dim Grade As Dom_Grade
    Set Grade = This.Factory.CreateGrade(Row.Grade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.Factory.CreateClassNo(Row.ClassNo)
    Dim Gender As Dom_Gender
    Set Gender = This.Factory.CreateGender(Row.Gender)
    Set CreateID = This.Factory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
End Function
