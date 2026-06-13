VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PeriodFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Formatter")
Option Explicit

Public Function Format(ByVal KeyCode As App_ClassHourKeyCode, ByVal ClassHourRMs As App_ClassHourReadModels, ByVal MainStreamRM As App_MainStreamReadModel) As Variant()
    Dim Resolver As App_ClassHourTypeModelResolver
    Set Resolver = New App_ClassHourTypeModelResolver
    Dim Key As String
    Key = Resolver.ParseKey(KeyCode)
    Dim GradeCount As Long
    GradeCount = MainStreamRM.UpperGrade
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Long
    Body = BuildBody(ClassHourRMs, Key, GradeCount)
    Format = BuildTable(Header, Body)
End Function

Private Function BuildHeader(ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount + 1)
    Dim C As Long
    For C = 2 To GradeCount + 1
        Result(1, C) = C - 1 & "”N"
    Next
    BuildHeader = Result
End Function

Private Function BuildBody(ByVal ClassHourRMs As App_ClassHourReadModels, ByVal Key As String, ByVal GradeCount As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To 1, 1 To GradeCount)
    Dim i As Long
    For i = 1 To ClassHourRMs.TimeTableCount
        Dim Model As App_TimeTableReadModel
        Set Model = ClassHourRMs.TimeTableItem(i)
        If Model.Key = Key Then
            If 0 < Model.Period Then
                If Model.Mark <> vbNullString Then
                    Result(1, Model.Grade) = Result(1, Model.Grade) + 1
                End If
            End If
        End If
    Next
    BuildBody = Result
End Function

Private Function BuildTable(ByRef Header() As Variant, ByRef Body() As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To UBound(Header) + UBound(Body), LBound(Header, 2) To UBound(Header, 2))
    Dim R As Long
    For R = 1 To UBound(Header)
        Dim C As Long
        For C = LBound(Header, 2) To UBound(Header, 2)
            Result(R, C) = Header(R, C)
        Next
    Next
    For R = UBound(Header) + 1 To UBound(Body)
        For C = LBound(Header, 2) To UBound(Header, 2)
            Result(R, C) = Body(R, C)
        Next
    Next
    BuildTable = Result
End Function
