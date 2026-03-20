Attribute VB_Name = "Dom_Enum"
'@Folder "Domain.Enum"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    domerr = 1000
End Enum

Public Enum Dom_ErrNum
    DomErrCanNotParse = vbObjectError + domerr
    DomErrEmptyDate
    DomErrEmptyObject
    DomErrInvalidDateFormat
    DomErrInvalidDateValue
    DomErrInvalidNaming
    DomErrInvalidRange
    DomErrInvalidTypeOfObject
    DomErrInvalidValue
    DomErrNotBoolean
    DomErrNotExistsItem
    DomErrNothingObject
    DomErrNotInteger
    DomErrNotNumeric
    DomErrNullString
End Enum

