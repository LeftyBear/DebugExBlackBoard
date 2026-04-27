VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SubjectCatalogFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    Factory As App_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As App_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByRef RawRows() As String) As Dom_SubjectCatalog
    Dim Result As Dom_SubjectCatalog
    Set Result = New Dom_SubjectCatalog
    Dim R As Long
    For R = LBound(RawRows) + 1 To UBound(RawRows)
        Dim C As Long
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            Select Case RawRows(1, C)
            Case jpnSyukeiMei
                Dim Name As String
                Name = RawRows(R, C)
            Case jpnTaisyoGakunen
                Dim TargetGrade As String
                TargetGrade = RawRows(R, C)
            Case jpnKigou
                Dim Mark As String
                Mark = RawRows(R, C)
            End Select
        Next
        Dim Subject As Dom_Subject
        Set Subject = This.Factory.CreateSubject(R - 1, Name, TargetGrade, Mark)
        Result.Add Subject
    Next
    Set Create = Result
End Function
