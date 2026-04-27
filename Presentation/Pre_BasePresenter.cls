VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_BasePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Presenter")
Option Explicit

Private Type Member
    Factory As Pre_PresenterResultFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Pre_PresenterResultFactory)
    Set This.Factory = Factory
End Sub

Public Sub Execute(ByVal Callback As Pre_IPresenterCallback)
    On Error GoTo ErrHandler
    Dim Factory As New Pre_PresenterResultFactory
    Callback.ExecuteCore
    Dim Result As Pre_PresenterResult
    Set Result = This.Factory.CreateSuccess
    GoTo Finally
ErrHandler:
    If Pre_ErrorUtility.IsDomainError(Err.Number) Or Pre_ErrorUtility.IsInfrastructureError(Err.Number) Then
        Dim Message As String
        Message = "î‘çÜ: " & Err.Number & vbCrLf & _
                  "èÍèä: " & Err.Source & vbCrLf & _
                  "ê‡ñæ: " & Err.Description
        Set Result = This.Factory.CreateBusinessError(Message)
    Else
        Callback.LogSystemError Err
        Set Result = This.Factory.CreateSystemError
    End If
Finally:
    Callback.RenderResult Result
End Sub
