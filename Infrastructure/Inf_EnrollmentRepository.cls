VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit

Implements Dom_IEnrollmentRepository

Private Type Member
    SourceData As Variant
    LastRow As Long
End Type

Private This As Member

Public Sub Initialize(ByVal SourceData As Variant)
    This.SourceData = SourceData
    This.LastRow = UBound(SourceData, 1)
End Sub

Private Function Dom_IEnrollmentRepository_LoadAggregate(ByVal TargetDate As Date) As Dom_EnrollmentAggregate
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Aggregate.Initialize
    Dim RowIndex As Long
    For RowIndex = 1 To This.LastRow
        If IsTargetRow(RowIndex, TargetDate) Then
            Dim StreamType As Dom_StreamType
            Dim Grade As Long
            Dim ClassNo As Long
            Dim Gender As Dom_Gender
            Dim Count As Long
            StreamType = ParseStreamType(RowIndex)
            Grade = ParseGrade(RowIndex)
            ClassNo = ParseClassNo(RowIndex)
            Gender = ParseGender(RowIndex)
            Count = ParseCount(RowIndex)
            If 0 < Count Then
                Dim Rec As Dom_EnrollmentRecord
                Set Rec = New Dom_EnrollmentRecord
                Rec.Initialize StreamType, Grade, ClassNo, Gender, Count
                Aggregate.Add Rec
            End If
        End If
    Next
    Set Dom_IEnrollmentRepository_LoadAggregate = Aggregate
End Function

Private Function IsTargetRow(ByVal RowIndex As Long, ByVal TargetDate As Date) As Boolean
    Dim RowDate As Date
    RowDate = This.SourceData(RowIndex, 1)
    If RowDate = TargetDate Then
        IsTargetRow = True
    End If
End Function

Private Function ParseStreamType(ByVal RowIndex As Long) As Dom_StreamType
    Dim Value As String
    Value = CStr(This.SourceData(RowIndex, 2))
    If Value = "Main" Then
        ParseStreamType = Dom_StreamType.Main
    Else
        ParseStreamType = Dom_StreamType.Special
    End If
End Function

Private Function ParseGrade(ByVal RowIndex As Long) As Long
    ParseGrade = CLng(This.SourceData(RowIndex, 3))
End Function

Private Function ParseClassNo(ByVal RowIndex As Long) As Long
    ParseClassNo = CLng(This.SourceData(RowIndex, 4))
End Function

Private Function ParseGender(ByVal RowIndex As Long) As Dom_Gender
    Dim Value As String
    Value = CStr(This.SourceData(RowIndex, 5))
    If Value = "Male" Then
        ParseGender = Dom_Gender.Male
    Else
        ParseGender = Dom_Gender.Female
    End If
End Function

Private Function ParseCount(ByVal RowIndex As Long) As Long
    ParseCount = CLng(This.SourceData(RowIndex, 6))
End Function
