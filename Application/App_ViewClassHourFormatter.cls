VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewClassHourFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Formatter")
Option Explicit

Private Type Member
    Calculator  As Dom_ClassHourCalculator
    Factory     As App_ValueObjectFactory
    Service     As App_ClassHourService
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Calculator = New Dom_ClassHourCalculator
    Set This.Factory = New App_ValueObjectFactory
    Set This.Service = New App_ClassHourService
End Sub

Public Function Format(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim ClassHourTable() As Long
    ClassHourTable = ToClassHourTable(ClassHourTypeEnum, ClassHour, TargetDate, Structure)
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

Private Function ToClassHourTable(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Structure As Dom_SchoolStructure) As Long()
    Dim Table() As Long
    ReDim Table(1 To Structure.SubjectCount + 1, 1 To Structure.MainStreamGradeCount)
    If ClassHour Is Nothing Then GoTo Finalize
    Dim Total() As Long
    ReDim Total(1 To Structure.MainStreamGradeCount)
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Factory.CreateClassHourType(ClassHourTypeEnum)
    Dim G As Long
    For G = 1 To Structure.MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = This.Factory.CreateGrade(G)
        Dim i As Long
        For i = 1 To Structure.SubjectCount
            Dim Subject As Dom_Subject
            Set Subject = Structure.SubjectIndexOf(i)
            Dim Filter As Dom_ClassHourFilter
            Set Filter = CreateClassHourFilter(ClassHourType, Subject, Grade)
            Dim Value As Long
            Value = This.Service.GetDailyHourTotal(ClassHour, TargetDate, This.Calculator, Filter)
            Table(i, G) = Value
            Total(G) = Total(G) + Value
        Next
        Table(UBound(Table), G) = Total(G)
    Next
Finalize:
    ToClassHourTable = Table
End Function

Private Function CreateClassHourFilter(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    With New Dom_ClassHourFilter
        Set CreateClassHourFilter = .Create(ClassHourType, Subject, Grade)
    End With
End Function
