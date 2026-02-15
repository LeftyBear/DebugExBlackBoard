Attribute VB_Name = "App_ErrorHandler"
'@Folder "Application.Handler.Error"
Option Explicit
Option Private Module

Public Function Handle(ByVal ErrNumber As Long, ByVal Source As String, ByVal Message As String) As App_ErrorPresenterResult
    Dim Code As App_ViewResult
    Code = TranslateError(ErrNumber)
    Dim Result As App_ErrorPresenterResult
    Set Result = New App_ErrorPresenterResult
'    Result.Initialize Code, Source, Message
    Set Handle = Result
End Function

Private Function TranslateError(ByVal ErrNumber As Long) As App_ViewResult
    Select Case True
        Case (ErrNumber = 0)
            TranslateError = App_ViewResult.Success
        Case IsDomainError(ErrNumber)
            TranslateError = App_ViewResult.BusinessError
        Case Else
            TranslateError = App_ViewResult.SystemError
    End Select
End Function

Private Function IsDomainError(ByVal ErrNumber As Long) As Boolean
    Dim Base As Long
    Base = ErrNumber - vbObjectError
    IsDomainError = (Dom_LayerErrNum.DomErr <= Base And Base < App_LayerErrNum.AppErr)
End Function
