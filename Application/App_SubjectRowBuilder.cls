VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SubjectRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Public Function Build(ByVal Records As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim HeaderRows As VBA.Collection
    Set HeaderRows = Records.Item(1)
    Dim R As Long
    For R = 2 To Records.Count
        Dim Rows As VBA.Collection
        Set Rows = Records.Item(R)
        Dim Row As App_SubjectRow
        Set Row = New App_SubjectRow
        Dim C As Long
        For C = 1 To Rows.Count
            If HeaderRows.Item(C) = jpnSyukeiMei Then
                Row.Name = Rows.Item(C)
            ElseIf HeaderRows.Item(C) = jpnTaisyoGakunen Then
                Row.TargetGrade = Rows.Item(C)
            ElseIf HeaderRows.Item(C) = jpnKigou Then
                Row.Mark = Rows.Item(C)
            End If
        Next
        Result.Add Row
    Next
    Set Build = Result
End Function
