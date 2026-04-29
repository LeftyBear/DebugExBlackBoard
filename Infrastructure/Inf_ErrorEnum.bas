Attribute VB_Name = "Inf_ErrorEnum"
'@Folder("Infrastructure.ValueObject")
Option Explicit
Option Private Module

Public Enum Inf_LayerErrNum
    InfErr = 4000
End Enum

Public Enum Inf_ErrNum
    InfErrNotFoundFile = vbObjectError + InfErr
    InfErrNotFoundLayerPrefix
    InfErrUnsupportedComponentType
End Enum
