Attribute VB_Name = "Pre_ErrorUtility"
'@Folder("Presentation.Utility")
Option Explicit
Option Private Module

Public Enum Pre_LayerErrNum
    PreErr = 3000
End Enum

Public Enum Pre_ErrNum
    PreErrSomothing = vbObjectError + PreErr
End Enum

Public Function IsDomainError(ByVal ErrNumber As Long) As Boolean
    Dim BaseNumber As Long
    BaseNumber = ErrNumber - vbObjectError
    IsDomainError = (Dom_LayerErrNum.DomErr <= BaseNumber And BaseNumber < App_LayerErrNum.AppErr)
End Function

Public Function IsInfrastructureError(ByVal ErrNumber As Long) As Boolean
    Dim BaseNumber As Long
    BaseNumber = ErrNumber - vbObjectError
    IsInfrastructureError = (Inf_LayerErrNum.InfErr <= BaseNumber And BaseNumber < Inf_LayerErrNum.InfErr + 100)
End Function

