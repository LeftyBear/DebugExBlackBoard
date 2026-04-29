Attribute VB_Name = "Dom_ErrorEnum"
'@Folder("Domain.ValueObject")
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    DomErr = 1000
End Enum

Public Enum Dom_ErrNum
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
