Attribute VB_Name = "Dom_Enum"
'@Folder "Domain.Enum"
Option Explicit
Option Private Module

Public Enum Dom_LayerErrNum
    DomErr = 1000
End Enum

Public Enum Dom_ErrNum
    errEmptyColumns = vbObjectError + DomErr
    errEmptyHeaderName
    errEmptyRows
    errInvalidColumnType
    errInvalidGradeValue
    errInvalidHeaderFormat
    errInvalidRowType
    errNothingClassColumn
    errNothingColumns
    errNothingGenderColumn
    errNothingGradeColumn
    errNothingRows
End Enum

