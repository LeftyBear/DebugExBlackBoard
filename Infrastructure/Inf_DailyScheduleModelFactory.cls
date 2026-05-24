VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyScheduleModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Type Member
    TypeMapper      As Inf_StreamTypeMapper
    GradeMapper     As Inf_GradeMapper
    ClassNoMapper   As Inf_ClassNoMapper
    GenderMapper    As Inf_GenderMapper
    TransferMapper  As Inf_TransferMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.TypeMapper = New Inf_StreamTypeMapper
    Set This.GradeMapper = New Inf_GradeMapper
    Set This.ClassNoMapper = New Inf_ClassNoMapper
    Set This.GenderMapper = New Inf_GenderMapper
    Set This.TransferMapper = New Inf_TransferMapper
    Dim Factory As Dom_ValueObjectFactory
    Set Factory = New Dom_ValueObjectFactory
    This.TypeMapper.Initialize Factory
    This.GradeMapper.Initialize Factory
    This.ClassNoMapper.Initialize Factory
    This.GenderMapper.Initialize Factory
    This.TransferMapper.Initialize Factory
End Sub

Public Function Create(ByVal DailyRows As Inf_ScheduleRows) As App_ScheduleReadModels
    Dim Result As App_ScheduleReadModels
    Set Result = New App_ScheduleReadModels
    Dim i As Long
    For i = 1 To DailyRows.Count
        Dim Row As Inf_ScheduleRow
        Set Row = DailyRows.Item(i)
        Dim Model As App_ScheduleReadModel
        Set Model = New App_ScheduleReadModel
        If Row.RawID <> vbNullString Then
            Model.RawID = Row.RawID
        ElseIf Row.Value <> vbNullString Then
            Model.Value = Row.Value
        End If
        Result.Add Model
    Next
    Set Create = Result
End Function
