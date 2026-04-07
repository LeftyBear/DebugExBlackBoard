Attribute VB_Name = "Dom_ErrorModel"
'@Folder "Domain.Model"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    domerr = 1000
End Enum

Public Enum Dom_ErrNum
    DomErrCanNotParse = vbObjectError + domerr
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

