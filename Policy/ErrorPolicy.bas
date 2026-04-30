Attribute VB_Name = "ErrorPolicy"
'@Folder("Policy")
Option Explicit
Option Private Module

Public Enum DomLayerErrNum
    DomErr = 1000
End Enum

Public Enum AppLayerErrNum
    AppErr = 2000
End Enum

Public Enum PreLayerErrNum
    PreErr = 3000
End Enum

Public Enum InfLayerErrNum
    InfErr = 4000
End Enum

Public Enum DomErrNum
    DomErrCanNotParse = vbObjectError + DomErr
    DomErrEmptyDate
    DomErrEmptyFilter
    DomErrEmptyObject
    DomErrInvalidDateFormat
    DomErrInvalidDateValue
    DomErrInvalidNaming
    DomErrInvalidRange
    DomErrInvalidTypeOfObject
    DomErrInvalidValue
    DomErrNegativeNumber
    DomErrNotBoolean
    DomErrNotFourDigitNumber
    DomErrNotExistsItem
    DomErrNothingObject
    DomErrNotInteger
    DomErrNotNumeric
    DomErrNotUnique
    DomErrNullString
    DomErrUnmatch
End Enum

Public Enum AppErrNum
    AppErrEmptyArray = vbObjectError + AppErr
    AppErrInvalidFilePath
    AppErrNothingItem
End Enum

Public Enum PreErrNum
    PreErrSomothing = vbObjectError + PreErr
End Enum

Public Enum InfErrNum
    InfErrNotFoundFile = vbObjectError + InfErr
    InfErrNotFoundLayerPrefix
    InfErrUnsupportedComponentType
End Enum

Public Function IsDomainError(ByVal ErrNumber As Long) As Boolean
    Dim BaseNumber As Long
    BaseNumber = ErrNumber - vbObjectError
    IsDomainError = (DomLayerErrNum.DomErr <= BaseNumber And BaseNumber < AppLayerErrNum.AppErr)
End Function

Public Function IsInfrastructureError(ByVal ErrNumber As Long) As Boolean
    Dim BaseNumber As Long
    BaseNumber = ErrNumber - vbObjectError
    IsInfrastructureError = (InfLayerErrNum.InfErr <= BaseNumber And BaseNumber < InfLayerErrNum.InfErr + 100)
End Function
