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

Public Function Format(ByVal ClassHourType As Dom_ClassHourType, ByVal Record As Dom_ClassHourRecord, ByVal Config As Dom_SchoolConfig) As Variant()
    Dim ClassHourTable() As Long
    ClassHourTable = ToClassHourTable(Plan, Record, Config)
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
                OutputTable(R, C) = Config.SubjectName(R - 1)
            ElseIf R = UBound(ClassHourTable, 1) + 1 And C = 1 Then
                OutputTable(R, C) = jpnGoukei
            Else
                OutputTable(R, C) = ClassHourTable(R - 1, C - 1)
            End If
        Next
    Next
    Format = OutputTable
End Function

Private Function ToClassHourTable(ByVal ClassHourType As Dom_ClassHourType, ByVal Record As Dom_ClassHourRecord, ByVal Config As Dom_SchoolConfig) As Long()
    Dim Table() As Long
    ReDim Table(1 To Config.SubjectCount + 1, 1 To Config.MainStreamGradeCount)
    Dim Total() As Long
    ReDim Total(1 To Config.MainStreamGradeCount)
    Dim Grade As Long
    For Grade = 1 To Config.MainStreamGradeCount
        Dim SubjectIndex As Long
        For SubjectIndex = 1 To Config.SubjectCount
            Dim Value As Long
            Value = Record.Value(ClassHourType, SubjectIndex, Grade)
            Table(SubjectIndex, Grade) = Value
            Total(Grade) = Total(Grade) + Value
        Next
        Table(UBound(Table), Grade) = Total(Grade)
    Next
    ToClassHourTable = Table
End Function
