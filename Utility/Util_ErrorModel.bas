Attribute VB_Name = "Util_ErrorModel"
'@Folder "Utility.Enum"
Option Explicit
Option Private Module

Public Enum Util_LayerErrNum
    UtilErr = 4000
End Enum

Public Enum Util_ErrNum
    UtilErrNotFoundLayerPrefix = vbObjectError + UtilErr
    UtilErrUnsupportedComponentType
End Enum

