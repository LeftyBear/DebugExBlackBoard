VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailyPeriodPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyPeriodPresenter

Public Function App_IDailyPeriodPresenter_CreateViewModel(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal MainStreamReadModel As App_MainStreamReadModel) As Variant()
    App_IDailyPeriodPresenter_CreateViewModel = Format(ExeKeyCode, ClassHourReadModels, MainStreamReadModel)
End Function

Private Function Format(ByVal KeyCode As App_ClassHourKeyCode, ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal MainStreamReadModel As App_MainStreamReadModel) As Variant()
    Dim Resolver As App_ClassHourTypeModelResolver
    Set Resolver = New App_ClassHourTypeModelResolver
    Dim Key As String
    Key = Resolver.ParseKey(KeyCode)
    Dim GradeCount As Long
    GradeCount = MainStreamReadModel.UpperGrade
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Long
    Body = BuildBody(ClassHourReadModels, Key, GradeCount)
    Format = BuildTable(Header, Body)
End Function

Private Function BuildHeader(ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount)
    Dim C As Long
    For C = 1 To GradeCount
        Result(1, C) = C & "”N"
    Next
    BuildHeader = Result
End Function

Private Function BuildBody(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal Key As String, ByVal GradeCount As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To 1, 1 To GradeCount)
    If ClassHourReadModels Is Nothing Then BuildBody = Result: Exit Function
    Dim i As Long
    For i = 1 To ClassHourReadModels.TimeTableCount
        Dim Model As App_TimeTableReadModel
        Set Model = ClassHourReadModels.TimeTableItem(i)
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
