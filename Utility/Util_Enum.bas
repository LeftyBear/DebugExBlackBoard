Attribute VB_Name = "Util_Enum"
'@Folder "Utility.Enum"
Option Explicit
Option Private Module

Public Enum Util_Direction
    D1 = 1
    D2
End Enum

Public Enum Util_LayerErrNum
    UtilErr = 4000
End Enum

Public Enum Util_ErrNum
    NotBoolean = vbObjectError + UtilErr
    NotFoundLayerPrefix
    NotInteger
    NotNumeric
    EmptyDate
    InvalidDateFormat
    InvalidDateValue
    UnsupportedComponentType
End Enum

