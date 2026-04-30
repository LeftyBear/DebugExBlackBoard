VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_IPresenterCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Interface")
Option Explicit

Public Sub ExecuteCore()
End Sub

Public Sub RenderResult(ByVal Result As Pre_PresenterResult)
End Sub

Public Sub LogSystemError(ByVal Error As ErrObject)
End Sub
