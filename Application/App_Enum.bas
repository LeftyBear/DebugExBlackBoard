Attribute VB_Name = "App_Enum"
'@Folder "Application.Enum"
Option Explicit
Option Private Module

Public Enum App_LayerErrNum
    AppErr = 2000
End Enum

Public Enum App_ErrNum
    InvalidRange = vbObjectError + AppErr
    InvalidHeaderName
    NotDefinedStructure
    NotFoundKey
    NotFoundSection
    NotPositiveNumber
    UnKnownHeaderName
End Enum

Public Enum App_ViewResult
    Success
    BusinessError
    SystemError
End Enum

