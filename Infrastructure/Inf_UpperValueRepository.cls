VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_UpperValueRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_IUpperValueRepository
Private Type Member
    Persistence As Inf_UpperValuePersistence
    EntList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal Persistence As Inf_UpperValuePersistence)
    Set This.Persistence = Persistence
    Set This.EntList = New Scripting.Dictionary
End Sub

Private Sub Dom_IUpperValueRepository_Add(ByVal Entry As Dom_UpperValue)
    This.EntList.Add "Only", Entry
End Sub

Private Function Dom_IUpperValueRepository_Item() As Dom_UpperValue
    Set Dom_IUpperValueRepository_Item = This.EntList.Item("Only")
End Function

Private Function Dom_IUpperValueRepository_ReadAll() As VBA.Collection
    Set Dom_IUpperValueRepository_ReadAll = This.Persistence.ReadAll
End Function

Private Function Dom_IUpperValueRepository_Replace(ByVal NewEntry As Dom_UpperValue) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Set This.EntList.Item("Only") = NewEntry
    Set Dom_IUpperValueRepository_Replace = Result
End Function
