Attribute VB_Name = "App_PresenterResultFactory"
'@Folder "Application.Factory"
Option Explicit

Public Function CreateSuccess() As App_PresenterResult
    Dim Result As App_PresenterResult
    Set Result = New App_PresenterResult
    Result.Initialize Success, vbNullString
    Set CreateSuccess = Result
End Function

Public Function CreateBusinessError(ByVal Message As String) As App_PresenterResult
    Dim Result As App_PresenterResult
    Set Result = New App_PresenterResult
    Result.Initialize BusinessError, Message
    Set CreateBusinessError = Result
End Function

Public Function CreateSystemError() As App_PresenterResult
    Dim Result As App_PresenterResult
    Set Result = New App_PresenterResult
    Result.Initialize SystemError, vbNullString
    Set CreateSystemError = Result
End Function
