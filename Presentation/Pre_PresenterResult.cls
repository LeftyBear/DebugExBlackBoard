VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_PresenterResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Presenter"
Option Explicit

Private Type Member
    ResultType  As Pre_ViewResultTypeEnum
    Message     As String
End Type

Private This As Member

Public Sub Initialize(ByVal ResultType As Pre_ViewResultTypeEnum, ByVal Message As String)
    This.ResultType = ResultType
    This.Message = Message
End Sub

Public Property Get ResultType() As Pre_ViewResultTypeEnum
    ResultType = This.ResultType
End Property

Public Property Get Message() As String
    Message = This.Message
End Property
