VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_MainStreamSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type TMember
    LowerGrade As Long
    UpperGrade As Long
    LowerClass As Long
    UpperClass As Long
    MaleGradeTotal() As Long
    FemaleGradeTotal() As Long
    ClassTotal() As Long
    MaleGrandTotal As Long
    FemaleGrandTotal As Long
    IsBuilt As Boolean
End Type

Private This As TMember

Friend Sub Build(ByVal Entities As VBA.Collection)
    If This.IsBuilt Then Err.Raise Dom_ErrNum.EmptyEntities, , "Class is built"
    Dim First As Dom_Enrollment
    Set First = Entities.Item(1)
    This.LowerGrade = First.MainStreamLowerGrade
    This.UpperGrade = First.MainStreamUpperGrade
    This.LowerClass = First.MainStreamLowerClass
    This.UpperClass = First.MainStreamUpperClass
    ReDim This.MaleGradeTotal(This.LowerGrade To This.UpperGrade)
    ReDim This.FemaleGradeTotal(This.LowerGrade To This.UpperGrade)
    ReDim This.ClassTotal(This.LowerGrade To This.UpperGrade, This.LowerClass To This.UpperClass)
    Dim Entity As Dom_Enrollment
    For Each Entity In Entities
        Dim Grade As Long
        For Grade = This.LowerGrade To This.UpperGrade
            Dim MaleTotal As Long
            MaleTotal = Entity.SumMainStreamMale(Grade)
            Dim FemaleTotal As Long
            FemaleTotal = Entity.SumMainStreamFemale(Grade)
            This.MaleGradeTotal(Grade) = This.MaleGradeTotal(Grade) + MaleTotal
            This.FemaleGradeTotal(Grade) = This.FemaleGradeTotal(Grade) + FemaleTotal
            This.MaleGrandTotal = This.MaleGrandTotal + MaleTotal
            This.FemaleGrandTotal = This.FemaleGrandTotal + FemaleTotal
            Dim ClassNo As Long
            For ClassNo = This.LowerClass To This.UpperClass
                This.ClassTotal(Grade, ClassNo) = This.ClassTotal(Grade, ClassNo) + Entity.SumMainStreamClass(Grade, ClassNo)
            Next
        Next
    Next
    This.IsBuilt = True
End Sub

Public Property Get MaleGrandTotal() As Long
    MaleGrandTotal = This.MaleGrandTotal
End Property

Public Property Get FemaleGrandTotal() As Long
    FemaleGrandTotal = This.FemaleGrandTotal
End Property

Public Function MaleGradeTotal(ByVal Grade As Long) As Long
    MaleGradeTotal = This.MaleGradeTotal(Grade)
End Function

Public Function FemaleGradeTotal(ByVal Grade As Long) As Long
    FemaleGradeTotal = This.FemaleGradeTotal(Grade)
End Function

Public Function ClassTotal(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ClassTotal = This.ClassTotal(Grade, ClassNo)
End Function
