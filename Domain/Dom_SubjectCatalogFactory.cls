VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SubjectCatalogFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByRef RawCSVRows() As String) As Dom_SubjectCatalog
    Dim Result As Dom_SubjectCatalog
    Set Result = New Dom_SubjectCatalog
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Select Case RawCSVRows(1, C)
            Case jpnSyukeiMei
                Dim Name As String
                Name = RawCSVRows(R, C)
            Case jpnTaisyoGakunen
                Dim TargetGrade As String
                TargetGrade = RawCSVRows(R, C)
            Case jpnKigou
                Dim Mark As String
                Mark = RawCSVRows(R, C)
            End Select
        Next
        With New Dom_Subject
            Dim Subject As Dom_Subject
            Set Subject = .Create(R - 1, Name, TargetGrade, Mark)
        End With
        Result.Add Subject
    Next
    Set Create = Result
End Function
