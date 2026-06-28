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
    Dim Resolver As App_StreamTypeResolver
    Set Resolver = New App_StreamTypeResolver
    Dim i As Long
    For i = 1 To Entries.Count
        Dim Model As App_EnrollmentReadModel
        Set Model = New App_EnrollmentReadModel
        Dim Entry As Dom_Enrollment
        Set Entry = Entries.Item(i)
        Model.TypeKey = Resolver.ParseKey(Entry.ID.StreamType.Value)
        Model.Grade = Entry.ID.Grade.Value
        Model.ClassNo = Entry.ID.ClassNo.Value
        Model.ClassName = Entry.ID.ClassName.Value
        Model.Gender = Entry.ID.Gender.Value
        Model.Value = Entry.Value.Value
        Model.Transfer = CStr(Entry.Transfer.IsTrue)
        Model.Remarks = Entry.Remarks.Value
        If Entry.ID.StreamType.Value = MainCode Then
            If Entry.ID.Gender.IsMale Then
                Result.AddMainMale Model
            ElseIf Entry.ID.Gender.IsFemale Then
                Result.AddMainFemale Model
            End If
        ElseIf Entry.ID.StreamType.Value = SpecialCode Then
            If Entry.ID.Gender.IsMale Then
                Result.AddSpecialMale Model
            ElseIf Entry.ID.Gender.IsFemale Then
                Result.AddSpecialFemale Model
            End If
        End If
    Next
    Set Create = Result
End Function
