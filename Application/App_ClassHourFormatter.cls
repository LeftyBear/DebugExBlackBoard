VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Calculator  As App_ClassHourSummaryCalculator
    Factory     As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Calculator As App_ClassHourSummaryCalculator, ByVal Factory As Dom_ValueObjectFactory)
    Set This.Calculator = Calculator
    Set This.Factory = Factory
End Sub

Public Function Format(ByVal TypeCode As Dom_ClassHourTypePolicy, ByVal Aggregate As Dom_ClassHourYearAggregate, ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim ClassHourTable() As Long
    ClassHourTable = ToClassHourTable(TypeCode, Aggregate, BeginIndex, EndIndex, Structure)
    Dim OutputTable() As Variant
    ReDim OutputTable(1 To UBound(ClassHourTable, 1) + 1, 1 To UBound(ClassHourTable, 2) + 1)
    Dim C As Long
    For C = 1 To UBound(ClassHourTable, 2) + 1
        If C = 1 Then
            OutputTable(1, C) = vbNullString
        Else
            OutputTable(1, C) = CStr(C - 1) & jpnNen
        End If
    Next
    Dim R As Long
    For R = 2 To UBound(ClassHourTable, 1) + 1
        For C = 1 To UBound(ClassHourTable, 2) + 1
            If R < UBound(ClassHourTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = Structure.SubjectNameOf(R - 1)
            ElseIf R = UBound(ClassHourTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = jpnGoukei
            Else
                OutputTable(R, C) = ClassHourTable(R - 1, C - 1)
            End If
        Next
    Next
    Format = OutputTable
End Function

Private Function ToClassHourTable(ByVal TypeCode As Dom_ClassHourTypePolicy, ByVal Aggregate As Dom_ClassHourYearAggregate, ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Structure As Dom_SchoolStructure) As Long()
    Dim Table() As Long
    ReDim Table(1 To Structure.SubjectCount + 1, 1 To Structure.MainStreamGradeCount)
    If Aggregate Is Nothing Then GoTo Finalize
    Dim Total() As Long
    ReDim Total(1 To Structure.MainStreamGradeCount)
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Factory.CreateClassHourType(TypeCode)
    Dim G As Long
    For G = 1 To Structure.MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = This.Factory.CreateGrade(G)
        Dim i As Long
        For i = 1 To Structure.SubjectCount
            Dim Subject As Dom_Subject
            Set Subject = Structure.SubjectByIndex(i)
            Dim Condition As App_ClassHourFilterCondition
            Set Condition = This.Factory.CreateClassHourFilterCondition(ClassHourType, Subject, Grade)
            Dim Value As Long
            Value = This.Calculator.SumOf(Aggregate, BeginIndex, EndIndex, Condition)
            Table(i, G) = Value
            Total(G) = Total(G) + Value
        Next
        Table(UBound(Table), G) = Total(G)
    Next
Finalize:
    ToClassHourTable = Table
End Function
