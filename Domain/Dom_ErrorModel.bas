Attribute VB_Name = "Dom_ErrorModel"
'@Folder "Domain.Model"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    DomErr = 1000
End Enum

Public Enum Dom_ErrNum
    DomErrCanNotParse = vbObjectError + DomErr
    DomErrEmptyDate
    DomErrEmptyObject
    DomErrInvalidDateFormat
    DomErrInvalidDateValue
    DomErrInvalidNaming
    DomErrInvalidRange
    DomErrInvalidTypeOfObject
    DomErrInvalidValue
    DomErrNegativeNumber
    DomErrNotBoolean
    DomErrNotExistsItem
    DomErrNothingObject
    DomErrNotInteger
    DomErrNotNumeric
    DomErrNotUnique
    DomErrNullString
    DomErrUnmatch
End Enum

