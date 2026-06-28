Attribute VB_Name = "Pre_ArrayUtility"
'@Folder("Infrastructure.Service")
Option Explicit
Option Private Module

Public Function MergeHorizontally(ByRef Base() As Variant, ByRef Another() As Variant) As Variant()
    Dim Result() As Variant
    ReDim Result(LBound(Base) To UBound(Base), LBound(Base, 2) To UBound(Base, 2) + UBound(Another, 2))
    Dim R As Long
    For R = LBound(Base) To UBound(Base)
        Dim C As Long
        For C = LBound(Base, 2) To UBound(Base, 2)
            Result(R, C) = Base(R, C)
        Next
    Next
    For R = LBound(Base) To UBound(Base)
        For C = UBound(Base, 2) + 1 To UBound(Base, 2) + UBound(Another, 2)
            Result(R, C) = Another(R, C)
        Next
    Next
    MergeHorizontally = Result
End Function

Public Function MergeVertically(ByRef Base() As Variant, ByRef Another() As Variant) As Variant()
    Dim Result() As Variant
    ReDim Result(LBound(Base) To UBound(Base) + UBound(Another), LBound(Base, 2) To UBound(Base, 2))
    Dim R As Long
    For R = LBound(Base) To UBound(Base)
        Dim C As Long
        For C = LBound(Base, 2) To UBound(Base, 2)
            Result(R, C) = Base(R, C)
        Next
    Next
    For R = UBound(Base) + 1 To UBound(Another)
        For C = LBound(Base, 2) To UBound(Base, 2)
            Result(R, C) = Another(R, C)
        Next
    Next
    MergeVertically = Result
End Function
