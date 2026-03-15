Attribute VB_Name = "Dom_Enum"
'@Folder "Domain.Enum"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    DomErr = 1000
End Enum

Public Enum Dom_ErrNum
    DomErrEmptyDate = vbObjectError + DomErr
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

