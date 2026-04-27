Attribute VB_Name = "App_ErrorEnum"
'@Folder("Application.Enum")
Option Explicit
Option Private Module

Public Enum App_LayerErrNum
    AppErr = 2000
End Enum

Public Enum App_ErrNum
    AppErrEmptyArray = vbObjectError + AppErr
    AppErrInvalidFilePath
    AppErrNothingItem
End Enum
