VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_Presenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Presenter"
Option Explicit

Private Type Member
    Logger      As Inf_ILogger
    MainView    As App_MainView
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As Inf_ILogger)
    Set This.Logger = Logger
End Sub

Public Sub AttachView(ByVal MainView As App_MainView)
    Set This.MainView = MainView
End Sub

