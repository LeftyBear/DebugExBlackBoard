VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_PeriodRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_IPeriodRepository
Private Type Member
    Persistence As Inf_PeriodPersistence
    Value       As Dom_Period
End Type

Private This As Member

Public Sub Inject(ByVal Persistence As Inf_PeriodPersistence)
    Set This.Persistence = Persistence
End Sub

Private Sub EnsureLoaded()
    If Not This.Value Is Nothing Then Exit Sub
    Dim Row As Inf_PeriodRow
    Set Row = This.Persistence.ReadAll
    Dim Period As Dom_Period
    Set Period = Dom_ValueObjectFactory.CreatePeriod(Row.Value)
    Set This.Value = Period
End Sub

Private Function Dom_IPeriodRepository_Item() As Dom_Period
    EnsureLoaded
    Set Dom_IPeriodRepository_Item = This.Value
End Function

Private Sub Dom_IPeriodRepository_Replace(ByVal NewValue As Dom_Period)
    EnsureLoaded
    Set This.Value = NewValue
End Sub

Private Sub Dom_IPeriodRepository_Save()
    EnsureLoaded
    Dim Row As Inf_PeriodRow
    Set Row = This.Persistence.ReadAll
    ReplaceRow Row
    This.Persistence.WriteAll Row
End Sub

Private Sub ReplaceRow(ByRef Row As Inf_PeriodRow)
    Row.Value = This.Value.Value
End Sub
