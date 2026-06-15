VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_TimeTableFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Service.Formatter")
Option Explicit

Public Function Format(ByVal Models As App_ClassHourReadModels, ByVal KeyCode As App_ClassHourTypeKeyCode, ByVal PeriodModel As App_PeriodReadModel, ByVal MainModel As App_MainStreamReadModel) As Variant()
    Dim Resolver As App_ClassHourTypeResolver
    Set Resolver = New App_ClassHourTypeResolver
    Dim TypeKey As String
    TypeKey = Resolver.ParseKey(KeyCode)
    Dim PeriodCount As Long
    PeriodCount = PeriodModel.Value
    Dim GradeCount As Long
    GradeCount = MainModel.UpperGrade
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim MarkTable() As Long
    MarkTable = BuildMarkTable(Models, TypeKey, PeriodCount, GradeCount)
    Dim CountedTable() As Long
    CountedTable = CountMark(MarkTable)
    Dim Body() As Variant
    Body = BuildBody(CountedTable)
    Format = Pre_ArrayUtility.MergeArray(Header, Body)
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

Private Function BuildBody(ByRef Table() As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(LBound(Table) To UBound(Table), LBound(Table, 2) To UBound(Table, 2) + 1)
    Dim R As Long
    For R = LBound(Table) To UBound(Table)
        Dim C As Long
        For C = 1 To 1
            If C = 1 Then
                Result(R, C) = 1
            Else
                Result(R, C) = Table(R, C)
            End If
        Next
    Next
    BuildBody = Result
End Function

Private Function BuildMarkTable(ByVal Models As App_ClassHourReadModels, ByVal TypeKey As String, ByVal PeriodCount As Long, ByVal GradeCount As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To PeriodCount, 1 To GradeCount)
    Dim R As Long
    For R = 1 To PeriodCount
        Dim i As Long
        For i = 1 To Models.TimeTableCount
            If Models.TimeTableItem(i).TypeKey = TypeKey Then
                If Models.TimeTableItem(i).Period = R Then
                    Result(R, Models.TimeTableItem(i).Grade) = Models.TimeTableItem(i).Mark
                End If
            End If
        Next
    Next
    BuildMarkTable = Result
End Function

Private Function CountMark(ByRef MarkTable() As Long) As Long()
    Dim Result() As Long
    ReDim Result(LBound(MarkTable) To UBound(MarkTable) + 1, LBound(MarkTable, 2) To UBound(MarkTable, 2))
    Dim Grade As Long
    For Grade = LBound(MarkTable, 2) To UBound(MarkTable, 2)
        Dim Period As Long
        For Period = LBound(MarkTable) To UBound(MarkTable)
            Result(Period, Grade) = MarkTable(Period, Grade)
            Dim Count As Long
            Count = Count + MarkTable(Period, Grade)
        Next
        Result(Period + 1, Grade) = Count
        Count = 0
    Next
    CountMark = Result
End Function
