VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Entries As Dom_DailyEnrollments) As App_EnrollmentReadModels
    Dim Result As App_EnrollmentReadModels
    Set Result = New App_EnrollmentReadModels
    Dim Resolver As App_StreamTypeModelResolver
    Set Resolver = New App_StreamTypeModelResolver
    Dim i As Long
    For i = 1 To Entries.Count
        Dim Model As App_EnrollmentReadModel
        Set Model = New App_EnrollmentReadModel
        Dim Entry As Dom_Enrollment
        Set Entry = Entries.Item(i)
        Model.TypeKey = Resolver.ToModel(Entry.ID.StreamType)
        Model.Grade = Entry.ID.Grade.Value
        Model.ClassNo = Entry.ID.ClassNo.Value
        Model.Gender = Entry.ID.Gender.Value
        Model.Value = Entry.Value.Value
        Model.Transfer = Entry.Transfer
        Model.Remarks = Entry.Remarks
        Result.Add Model
    Next
    Set Create = Result
End Function
