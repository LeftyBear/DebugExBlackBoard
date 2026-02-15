VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.VO"
Option Explicit
Private Type MainStreamMember
    LowerGrade          As Long
    UpperGrade          As Long
    MaleGradeTotal()    As Long
    FemaleGradeTotal()  As Long
    LowerClass          As Long
    UpperClass          As Long
    ClassTotal()        As Long
    MaleGrandTotal      As Long
    FemaleGrandTotal    As Long
End Type

Private Type SpecialMember
    LowerGrade          As Long
    UpperGrade          As Long
    MaleGradeTotal()    As Long
    FemaleGradeTotal()  As Long
    LowerClass          As Long
    UpperClass          As Long
    ClassTotal()        As Long
    MaleGrandTotal      As Long
    FemaleGrandTotal    As Long
End Type

Private Type Member
    Entities    As VBA.Collection
    MainStream  As MainStreamMember
    Special     As SpecialMember
End Type

Private This As Member

Friend Sub Inject(ByVal Entities As VBA.Collection)
    If Entities Is Nothing Then
        Err.Raise Dom_ErrNum.NothingEntities, , "Entities is Nothing"
    End If
    If Entities.Count = 0 Then
        Err.Raise Dom_ErrNum.EmptyEntities, , "Entities is empty"
    End If
    Set This.Entities = Entities
    Dim First As Dom_Enrollment
    Set First = Entities.Item(1)
    This.MainStream.LowerGrade = First.MainStreamLowerGrade
    This.MainStream.UpperGrade = First.MainStreamUpperGrade
    This.MainStream.LowerClass = First.MainStreamLowerClass
    This.MainStream.UpperClass = First.MainStreamUpperClass
    This.Special.LowerGrade = First.SpecialLowerGrade
    This.Special.UpperGrade = First.SpecialUpperGrade
    This.Special.LowerClass = First.SpecialLowerClass
    This.Special.UpperClass = First.SpecialUpperClass
    ReDim This.MainStream.MaleGradeTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade)
    ReDim This.MainStream.FemaleGradeTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade)
    ReDim This.MainStream.ClassTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade, First.MainStreamLowerClass To First.MainStreamUpperClass)
    ReDim This.Special.MaleGradeTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade)
    ReDim This.Special.FemaleGradeTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade)
    ReDim This.Special.ClassTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade, First.SpecialLowerClass To First.SpecialUpperClass)
    AggregateEntities
End Sub

Public Function MainStreamLowerGrade() As Long
    MainStreamLowerGrade = This.MainStream.LowerGrade
End Function

Public Function MainStreamUpperGrade() As Long
    MainStreamUpperGrade = This.MainStream.UpperGrade
End Function

Public Function MainStreamLowerClass() As Long
    MainStreamLowerClass = This.MainStream.LowerClass
End Function

Public Function MainStreamUpperClass() As Long
    MainStreamUpperClass = This.MainStream.UpperClass
End Function

Public Function MainStreamMaleGradeTotal(ByVal Grade As Long) As Long
    MainStreamMaleGradeTotal = This.MainStream.MaleGradeTotal(Grade)
End Function

Public Function MainStreamFemaleGradeTotal(ByVal Grade As Long) As Long
    MainStreamFemaleGradeTotal = This.MainStream.FemaleGradeTotal(Grade)
End Function

Public Function MainStreamClassTotal(ByVal Grade As Long, ByVal Class As Long) As Long
    MainStreamClassTotal = This.MainStream.ClassTotal(Grade, Class)
End Function

Public Function MainStreamMaleGrandTotal() As Long
    MainStreamMaleGrandTotal = This.MainStream.MaleGrandTotal
End Function

Public Function MainStreamFemaleGrandTotal() As Long
    MainStreamFemaleGrandTotal = This.MainStream.FemaleGrandTotal
End Function

Public Function SpecialLowerGrade() As Long
    SpecialLowerGrade = This.Special.LowerGrade
End Function

Public Function SpecialUpperGrade() As Long
    SpecialUpperGrade = This.Special.UpperGrade
End Function

Public Function SpecialLowerClass() As Long
    SpecialLowerClass = This.Special.LowerClass
End Function

Public Function SpecialUpperClass() As Long
    SpecialUpperClass = This.MainStream.UpperClass
End Function

Public Function SpecialMaleGradeTotal(ByVal Grade As Long) As Long
    SpecialMaleGradeTotal = This.Special.MaleGradeTotal(Grade)
End Function

Public Function SpecialFemaleGradeTotal(ByVal Grade As Long) As Long
    SpecialFemaleGradeTotal = This.Special.FemaleGradeTotal(Grade)
End Function

Public Function SpecialClassTotal(ByVal Grade As Long, ByVal Class As Long) As Long
    SpecialClassTotal = This.Special.ClassTotal(Grade, Class)
End Function

Public Function SpecialMaleGrandTotal() As Long
    SpecialMaleGrandTotal = This.Special.MaleGrandTotal
End Function

Public Function SpecialFemaleGrandTotal() As Long
    SpecialFemaleGrandTotal = This.Special.FemaleGrandTotal
End Function

Private Sub AggregateEntities()
    Dim Entity As Dom_Enrollment
    For Each Entity In This.Entities
        AggregateMainStream Entity
        AggregateSpecial Entity
    Next
End Sub

Private Sub AggregateMainStream(ByVal Entity As Dom_Enrollment)
    With This.MainStream
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            Dim MaleGradeTotal As Long
            MaleGradeTotal = Entity.SumMainStreamMale(Grade)
            Dim FemaleGradeTotal As Long
            FemaleGradeTotal = Entity.SumMainStreamFemale(Grade)
            .MaleGradeTotal(Grade) = MaleGradeTotal
            .FemaleGradeTotal(Grade) = FemaleGradeTotal
            .MaleGrandTotal = .MaleGrandTotal + MaleGradeTotal
            .FemaleGrandTotal = .FemaleGrandTotal + FemaleGradeTotal
            Dim ClassNo As Long
            For ClassNo = .LowerClass To .UpperClass
                Dim ClassTotal As Long
                ClassTotal = Entity.SumMainStreamClass(Grade, ClassNo)
                .ClassTotal(Grade, ClassNo) = ClassTotal
            Next
        Next
    End With
End Sub

Private Sub AggregateSpecial(ByVal Entity As Dom_Enrollment)
    With This.Special
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            Dim MaleGradeTotal As Long
            MaleGradeTotal = Entity.SumSpecialMale(Grade)
            Dim FemaleGradeTotal As Long
            FemaleGradeTotal = Entity.SumSpecialFemale(Grade)
            .MaleGradeTotal(Grade) = MaleGradeTotal
            .FemaleGradeTotal(Grade) = FemaleGradeTotal
            .MaleGrandTotal = .MaleGrandTotal + MaleGradeTotal
            .FemaleGrandTotal = .FemaleGrandTotal + FemaleGradeTotal
            Dim ClassNo As Long
            For ClassNo = .LowerClass To .UpperClass
                Dim ClassTotal As Long
                ClassTotal = Entity.SumSpecialClass(Grade, ClassNo)
                .ClassTotal(Grade, ClassNo) = ClassTotal
            Next
        Next
    End With
End Sub

