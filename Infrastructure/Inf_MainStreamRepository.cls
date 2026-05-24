VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_IMainStreamRepository
Private Type Member
    Persistence As Inf_MainStreamPersistence
    Value       As Dom_MainStreamValue
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_MainStreamPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Value Is Nothing Then Exit Sub
    Dim Rows As Inf_MainStreamRows
    Set Rows = This.Persistence.ReadAll
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_MainStreamRow
        Set Row = Rows.Item(i)
        If 0 < Row.UpperGrade Then
            Dim UpperGrade As Long
            UpperGrade = Row.UpperGrade
        ElseIf 0 < Row.UpperClassNo Then
            Dim UpperClassNo As Long
            UpperClassNo = Row.UpperClassNo
        End If
        Dim Value As Dom_MainStreamValue
        Set Value = New Dom_MainStreamValue
        Value.Initialize UpperGrade, UpperClassNo
    Next
    Set This.Value = Value
End Sub

Private Function Dom_IMainStreamRepository_Item() As Dom_MainStreamValue
    EnsureLoaded
    Set Dom_IMainStreamRepository_Item = This.Value
End Function

Private Sub Dom_IMainStreamRepository_Replace(ByVal NewValue As Dom_MainStreamValue)
    EnsureLoaded
    Set This.Value = NewValue
End Sub

Private Sub Dom_IMainStreamRepository_Save()
    EnsureLoaded
    Dim Rows As Inf_MainStreamRows
    Set Rows = This.Persistence.ReadAll
    ReplaceRow Rows
    This.Persistence.WriteAll Rows
End Sub

Private Sub ReplaceRow(ByRef Rows As Inf_EnrollmentRows)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_MainStreamRow
        Set Row = Rows.Item(i)
        If 0 < Row.UpperGrade Then
            Row.UpperGrade = This.Value.UpperGrade
        ElseIf 0 < Row.UpperClassNo Then
            Row.UpperClassNo = This.Value.UpperClassNo
        End If
    Next
End Sub
