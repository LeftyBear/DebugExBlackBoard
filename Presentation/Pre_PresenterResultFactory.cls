VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_PresenterResultFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Factory"
Option Explicit

Public Function CreateSuccess() As Pre_PresenterResult
    Dim Result As Pre_PresenterResult
    Set Result = New Pre_PresenterResult
    Result.Initialize Success, vbNullString
    Set CreateSuccess = Result
End Function

Public Function CreateBusinessError(ByVal Message As String) As Pre_PresenterResult
    Dim Result As Pre_PresenterResult
    Set Result = New Pre_PresenterResult
    Result.Initialize BusinessError, Message
    Set CreateBusinessError = Result
End Function

Public Function CreateSystemError() As Pre_PresenterResult
    Dim Result As Pre_PresenterResult
    Set Result = New Pre_PresenterResult
    Result.Initialize SystemError, vbNullString
    Set CreateSystemError = Result
End Function
