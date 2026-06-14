VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ScheduleViewModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModelCollection")
Option Explicit
Private Type Member
    Models As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Models = New VBA.Collection
End Sub

Public Sub Add(ByVal Model As Pre_ScheduleViewModel)
    This.Models.Add Model
End Sub

Public Function Count() As Long
    Count = This.Models.Count
End Function

Public Function Item(ByVal Index As Long) As Pre_ScheduleViewModel
    Set Item = This.Models.Item(Index)
End Function
