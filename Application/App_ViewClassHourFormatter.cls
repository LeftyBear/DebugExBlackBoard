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

Public Function Format(ByVal ClassHourType As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourColumn, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim ClassHourTable() As Long
    ClassHourTable = ToClassHourTable(ClassHourType, ClassHour, Structure)
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

Private Function ToClassHourTable(ByVal ClassHourType As Dom_ClassHourTypeEnum, ByVal ClassHour As Dom_ClassHourColumn, ByVal Structure As Dom_SchoolStructure) As Long()
    Dim Table() As Long
    ReDim Table(1 To Structure.SubjectCount + 1, 1 To Structure.MainStreamGradeCount)
    Dim Total() As Long
    ReDim Total(1 To Structure.MainStreamGradeCount)
    Dim G As Long
    For G = 1 To Structure.MainStreamGradeCount
        Dim Grade As Dom_Grade
        Set Grade = CreateGrade(G)
        Dim i As Long
        For i = 1 To Structure.SubjectCount
            Dim Subject As Dom_Subject
            Set Subject = CreateSubject(i)
            Dim Filter As Dom_ClassHourFilter
            Set Filter = CreateFilter(ClassHourType, Subject, Grade)
            Dim Value As Long
            Value = ClassHour.GetHourTotal(Filter)
            Table(i, G) = Value
            Total(G) = Total(G) + Value
        Next
        Table(UBound(Table), G) = Total(G)
    Next
    ToClassHourTable = Table
End Function

Private Function CreateFilter(ByVal ClassHourTypeEnum As Dom_ClassHourTypeEnum, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Dom_ClassHourFilter
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = CreateClassHourType(ClassHourTypeEnum)
    With New Dom_ClassHourFilter
        Set CreateFilter = .Create(ClassHourType, Subject, Grade)
    End With
End Function

Private Function CreateClassHourType(ByVal Value As Long) As Dom_ClassHourType
    With New Dom_ClassHourType
        Set CreateClassHourType = .Create(Value)
    End With
End Function

Private Function CreateSubject(ByVal Index As Long) As Dom_Subject
    With New Dom_Subject
        Set CreateSubject = .Create(Index, vbNullString, vbNullString, vbNullString)
    End With
End Function

Private Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function
