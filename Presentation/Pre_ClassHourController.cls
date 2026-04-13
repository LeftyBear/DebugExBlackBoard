VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_ClassHourController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Controller")
Option Explicit

Private Type Member
    Presenter As Pre_Presenter
End Type

Private This As Member

Public Sub Initialize(ByVal Presenter As Pre_Presenter)
    Set This.Presenter = Presenter
End Sub

Public Sub ChangeDate(ByVal SelectedDate As Date)
    This.Presenter.LoadAsync SelectedDate
End Sub
