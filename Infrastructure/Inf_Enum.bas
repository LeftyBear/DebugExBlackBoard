Attribute VB_Name = "Inf_Enum"
'@Folder "Infrastructure.Enum"
Option Explicit
Option Private Module

Public Enum Inf_LayerErrNum
    InfErr = 3000
End Enum

Public Enum Inf_ErrNum
    InfErrNotFoundFile = vbObjectError + InfErr
End Enum
