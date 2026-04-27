VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    EntityFactory   As App_EntityFactory
    VOFactory       As App_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal EntityFactory As App_EntityFactory, ByVal VOFactory As App_ValueObjectFactory)
    Set This.EntityFactory = EntityFactory
    Set This.VOFactory = VOFactory
End Sub

Public Function Create(ByVal Rows As VBA.Collection) As Dom_EnrollmentYearAggregate
    Dim Yearly As Dom_EnrollmentYearAggregate
    Set Yearly = New Dom_EnrollmentYearAggregate
    Dim Grouped As Scripting.Dictionary
    Set Grouped = GroupByDate(Rows)
    Dim Key As Variant
    For Each Key In Grouped.Keys
        Dim Daily As Dom_EnrollmentAggregate
        Set Daily = CreateDaily(Grouped.Item(Key))
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function

Private Function GroupByDate(ByVal Rows As VBA.Collection) As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_EnrollmentRow
        Set Row = Rows.Item(i)
        Dim Key As String
        Key = CStr(Row.NormDate)
        If Not Result.Exists(Key) Then
            Result.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = Result.Item(Key)
        Items.Add Row
    Next
    Set GroupByDate = Result
End Function

Private Function CreateDaily(ByVal Rows As VBA.Collection) As Dom_EnrollmentAggregate
    Dim Result As Dom_EnrollmentAggregate
    Set Result = New Dom_EnrollmentAggregate
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_EnrollmentRow
        Set Row = Rows.Item(i)
        Dim DateID As Dom_DateID
        Set DateID = This.VOFactory.CreateDateID(Row.NormDate)
        Result.SetDateID DateID
        Result.SetTransfer This.EntityFactory.CreateTransfer(DateID, Row.Transfer)
        Result.SetRemarks This.EntityFactory.CreateRemarks(DateID, Row.Remarks)
        Result.AddEnrollment CreateEnrollment(DateID, Row)
    Next
    Set CreateDaily = Result
End Function

Private Function CreateEnrollment(ByVal DateID As Dom_DateID, ByVal Row As App_EnrollmentRow) As Dom_Enrollment
    Dim ID As Dom_EnrollmentID
    Set ID = CreateEnrollmentID(Row)
    Dim Value As Dom_EnrollmentValue
    Set Value = This.VOFactory.CreateEnrollmentValue(Row.Value)
    Set CreateEnrollment = This.EntityFactory.CreateEnrollment(DateID, ID, Value)
End Function

Private Function CreateEnrollmentID(ByVal Row As App_EnrollmentRow) As Dom_EnrollmentID
    Dim StreamType As Dom_StreamType
    Set StreamType = This.VOFactory.CreateStreamType(Row.TypeCode)
    Dim Grade As Dom_Grade
    Set Grade = This.VOFactory.CreateGrade(Row.Grade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.VOFactory.CreateClassNo(Row.ClassNo)
    Dim Gender As Dom_Gender
    Set Gender = This.VOFactory.CreateGender(Row.Gender)
    Set CreateEnrollmentID = This.VOFactory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
End Function
