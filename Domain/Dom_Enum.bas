Attribute VB_Name = "Dom_Enum"
'@Folder "Domain.Enum"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    DomErr = 1000
End Enum

Public Enum Dom_ErrNum
    DomErrEmptyObject = vbObjectError + DomErr
    DomErrInvalidRange
    DomErrInvalidNaming
    DomErrInvalidTypeOfObject
    DomErrInvalidValue
    DomErrNothingObject
    DomErrNullString
    DomErrNotBoolean
    DomErrEmptyDate
    DomErrInvalidDateFormat
    DomErrInvalidDateValue
    DomErrNotNumeric
    DomErrNotInteger
End Enum

