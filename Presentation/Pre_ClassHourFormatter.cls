VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ClassHourFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Formatter")
Option Explicit

Public Function Format(ByVal Models As App_ClassHourReadModels, ByVal KeyCode As App_ClassHourKeyCode, ByVal SubjectModels As App_SubjectReadModels, ByVal MainModel As App_MainStreamReadModel) As Variant()
    Dim Resolver As App_ClassHourTypeModelResolver
    Set Resolver = New App_ClassHourTypeModelResolver
    Dim Key As String
    Key = Resolver.ParseKey(KeyCode)
    Dim GradeCount As Long
    GradeCount = MainModel.UpperGrade
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim ValueTable() As Long
    ValueTable = BuildValueTable(Models, Key, SubjectModels, GradeCount)
    Dim TotalizedTable() As Long
    TotalizedTable = TotalizeValue(Models, Key, ValueTable)
    Dim Body() As Variant
    Body = BuildBody(TotalizedTable, SubjectModels)
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

Private Function BuildBody(ByRef Table() As Long, ByVal SubjectModels As App_SubjectReadModels) As Variant()
    Dim Result() As Variant
    ReDim Result(LBound(Table) To UBound(Table), LBound(Table, 2) To UBound(Table, 2) + 1)
    Dim R As Long
    For R = LBound(Table) To UBound(Table)
        Dim C As Long
        For C = 1 To SubjectModels.Count
            If C = 1 Then
                Result(R, C) = SubjectModels.Item(C).Name
            Else
                Result(R, C) = Table(R, C)
            End If
        Next
    Next
    BuildBody = Result
End Function

Private Function BuildTable(ByRef Header() As Variant, ByRef Body() As Variant) As Variant()
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

Private Function BuildValueTable(ByVal Models As App_ClassHourReadModels, ByVal Key As String, ByVal SubjectModels As App_SubjectReadModels, ByVal GradeCount As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To SubjectModels.Count, 1 To GradeCount)
    Dim R As Long
    For R = 1 To SubjectModels.Count
        Dim i As Long
        For i = 1 To Models.ClassHourCount
            If Models.ClassHourItem(i).Key = Key Then
                If Models.ClassHourItem(i).Subject = SubjectModels.Item(R).Name Then
                    Result(R, Models.ClassHourItem(i).Grade) = Models.ClassHourItem(i).Value
                End If
            End If
        Next
    Next
    BuildValueTable = Result
End Function

Private Function TotalizeValue(ByVal Models As App_ClassHourReadModels, ByVal Key As String, ByRef ValueTable() As Long) As Long()
    Dim Result() As Long
    ReDim Result(LBound(ValueTable) To UBound(ValueTable) + 1, LBound(ValueTable, 2) To UBound(ValueTable, 2))
    Dim Calculator As App_ClassHourCalculator
    Set Calculator = New App_ClassHourCalculator
    Dim Grade As Long
    For Grade = LBound(ValueTable, 2) To UBound(ValueTable, 2)
        Result(UBound(ValueTable) + 1, Grade) = Calculator.CalculateTotal(Models, Key, Grade)
    Next
    TotalizeValue = Result
End Function
