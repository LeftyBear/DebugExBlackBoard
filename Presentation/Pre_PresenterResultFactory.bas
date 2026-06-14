Attribute VB_Name = "Pre_PresenterResultFactory"
'@Folder("Presentation.ViewModelFactory")
Option Explicit

Public Function CreateSuccess() As App_UseCaseResult
    Dim Result As App_UseCaseResult
    Set Result = New App_UseCaseResult
    Result.Initialize SuccessCode, vbNullString
    Set CreateSuccess = Result
End Function

Public Function CreateBusinessError(ByVal Message As String) As App_UseCaseResult
    Dim Result As App_UseCaseResult
    Set Result = New App_UseCaseResult
    Result.Initialize BusinessErrorCode, Message
    Set CreateBusinessError = Result
End Function

Public Function CreateSystemError() As App_UseCaseResult
    Dim Result As App_UseCaseResult
    Set Result = New App_UseCaseResult
    Result.Initialize SystemErrorCode, vbNullString
    Set CreateSystemError = Result
End Function
