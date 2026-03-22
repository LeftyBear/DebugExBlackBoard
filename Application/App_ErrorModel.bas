Attribute VB_Name = "App_ErrorModel"
'@Folder "Application.Enum"
Option Explicit
Option Private Module

Public Enum App_LayerErrNum
    AppErr = 2000
End Enum

Public Enum App_ErrNum
    AppErrEmptyData = vbObjectError + AppErr
End Enum

Public Enum App_ViewResultType
    Success
    BusinessError
    SystemError
End Enum

Public Function IsDomainError(ByVal ErrNumber As Long) As Boolean
    Dim BaseNumber As Long
    BaseNumber = ErrNumber - vbObjectError
    IsDomainError = (Dom_LayerErrNum.DomErr <= BaseNumber And BaseNumber < App_LayerErrNum.AppErr)
End Function
