VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_GradeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_GRADE As String = "”N"

Public Function ToRaw(ByVal Grade As Dom_Grade) As String
    ToRaw = CStr(Grade.Value) & COLUMN_GRADE
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Grade
    If VBA.InStr(1, Raw, COLUMN_GRADE) = 0 Then Exit Function
    Dim EndPosition As Long
    EndPosition = VBA.InStr(1, Raw, COLUMN_GRADE)
    Dim BeginPosition As Long
    BeginPosition = EndPosition - 1
    Dim GradeString As String
    GradeString = VBA.Mid$(Raw, BeginPosition, EndPosition)
    Dim Grade As Long
    Grade = Inf_NumericUtility.ExtractNumber(GradeString)
    Set ToDomain = Dom_ValueObjectFactory.CreateGrade(Grade)
End Function
