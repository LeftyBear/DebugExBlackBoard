VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollMatrixBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Builder"
Option Explicit

Public Function BuildMatrix(ByVal Values As VBA.Collection, ByVal GradeCount As Long, ByVal ClassCount As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To GradeCount, 1 To ClassCount)
    Dim Index As Long
    Index = 1
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To ClassCount
            Result(Grade, ClassNo) = Values.Item(Index)
            Index = Index + 1
        Next
    Next
    BuildMatrix = Result
End Function
