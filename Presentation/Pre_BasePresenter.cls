VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_BasePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit

Public Sub Execute(ByVal Callback As Pre_IPresenterCallback)
    On Error GoTo ErrHandler
    Callback.ExecuteCore
    Dim Result As Pre_PresenterResult
    Set Result = Pre_PresenterResultFactory.CreateSuccess
    GoTo Finally
ErrHandler:
    If ErrorPolicy.IsDomainError(Err.Number) Or ErrorPolicy.IsInfrastructureError(Err.Number) Then
        Dim Message As String
        Message = "î‘çÜ: " & Err.Number & vbCrLf & _
                  "èÍèä: " & Err.Source & vbCrLf & _
                  "ê‡ñæ: " & Err.Description
        Set Result = Pre_PresenterResultFactory.CreateBusinessError(Message)
    Else
        Callback.LogSystemError Err
        Set Result = Pre_PresenterResultFactory.CreateSystemError
    End If
Finally:
    Callback.RenderResult Result
End Sub
