VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.VO"
Option Explicit
Private Type ThisMember
    Items As VBA.Collection
End Type
Private This As ThisMember
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
Private MainStream As MainStreamMember
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
Private Special As SpecialMember

Friend Sub Inject(ByVal Members As VBA.Collection)
    If Members Is Nothing Then
        Err.Raise Dom_ErrNum.NothingMembers, , "Members is Nothing"
    End If
    If Members.Count = 0 Then
        Err.Raise Dom_ErrNum.EmptyMembers, , "Members is empty"
    End If
    Set This.Items = Members
    Dim First As Dom_EnrollMember
    Set First = Members.Item(1)
    MainStream.LowerGrade = First.MainStreamLowerGrade
    MainStream.UpperGrade = First.MainStreamUpperGrade
    MainStream.LowerClass = First.MainStreamLowerClass
    MainStream.UpperClass = First.MainStreamUpperClass
    Special.LowerGrade = First.SpecialLowerGrade
    Special.UpperGrade = First.SpecialUpperGrade
    Special.LowerClass = First.SpecialLowerClass
    Special.UpperClass = First.SpecialUpperClass
    ReDim MainStream.MaleGradeTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade)
    ReDim MainStream.FemaleGradeTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade)
    ReDim MainStream.ClassTotal(First.MainStreamLowerGrade To First.MainStreamUpperGrade, First.MainStreamLowerClass To First.MainStreamUpperClass)
    ReDim Special.MaleGradeTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade)
    ReDim Special.FemaleGradeTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade)
    ReDim Special.ClassTotal(First.MainStreamLowerGrade To First.SpecialUpperGrade, First.SpecialLowerClass To First.SpecialUpperClass)
    AggregateMembers
End Sub

Public Function MainStreamLowerGrade() As Long
    MainStreamLowerGrade = MainStream.LowerGrade
End Function

Public Function MainStreamUpperGrade() As Long
    MainStreamUpperGrade = MainStream.UpperGrade
End Function

Public Function MainStreamLowerClass() As Long
    MainStreamLowerClass = MainStream.LowerClass
End Function

Public Function MainStreamUpperClass() As Long
    MainStreamUpperClass = MainStream.UpperClass
End Function

Public Function MainStreamMaleGradeTotal(ByVal Grade As Long) As Long
    MainStreamMaleGradeTotal = MainStream.MaleGradeTotal(Grade)
End Function

Public Function MainStreamFemaleGradeTotal(ByVal Grade As Long) As Long
    MainStreamFemaleGradeTotal = MainStream.FemaleGradeTotal(Grade)
End Function

Public Function MainStreamClassTotal(ByVal Grade As Long, ByVal Class As Long) As Long
    MainStreamClassTotal = MainStream.ClassTotal(Grade, Class)
End Function

Public Function MainStreamMaleGrandTotal() As Long
    MainStreamMaleGrandTotal = MainStream.MaleGrandTotal
End Function

Public Function MainStreamFemaleGrandTotal() As Long
    MainStreamFemaleGrandTotal = MainStream.FemaleGrandTotal
End Function

Public Function SpecialLowerGrade() As Long
    SpecialLowerGrade = Special.LowerGrade
End Function

Public Function SpecialUpperGrade() As Long
    SpecialUpperGrade = Special.UpperGrade
End Function

Public Function SpecialLowerClass() As Long
    SpecialLowerClass = Special.LowerClass
End Function

Public Function SpecialUpperClass() As Long
    SpecialUpperClass = MainStream.UpperClass
End Function

Public Function SpecialMaleGradeTotal(ByVal Grade As Long) As Long
    SpecialMaleGradeTotal = Special.MaleGradeTotal(Grade)
End Function

Public Function SpecialFemaleGradeTotal(ByVal Grade As Long) As Long
    SpecialFemaleGradeTotal = Special.FemaleGradeTotal(Grade)
End Function

Public Function SpecialClassTotal(ByVal Grade As Long, ByVal Class As Long) As Long
    SpecialClassTotal = Special.ClassTotal(Grade, Class)
End Function

Public Function SpecialMaleGrandTotal() As Long
    SpecialMaleGrandTotal = Special.MaleGrandTotal
End Function

Public Function SpecialFemaleGrandTotal() As Long
    SpecialFemaleGrandTotal = Special.FemaleGrandTotal
End Function

Private Sub AggregateMembers()
    Dim Member As Dom_EnrollMember
    For Each Member In This.Items
        AggregateMainStream Member
        AggregateSpecial Member
    Next
End Sub

Private Sub AggregateMainStream(ByVal Member As Dom_EnrollMember)
    With MainStream
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            Dim MaleGradeTotal As Long
            MaleGradeTotal = Member.SumMainStreamMale(Grade)
            Dim FemaleGradeTotal As Long
            FemaleGradeTotal = Member.SumMainStreamFemale(Grade)
            .MaleGradeTotal(Grade) = MaleGradeTotal
            .FemaleGradeTotal(Grade) = FemaleGradeTotal
            .MaleGrandTotal = .MaleGrandTotal + MaleGradeTotal
            .FemaleGrandTotal = .FemaleGrandTotal + FemaleGradeTotal
            Dim ClassNo As Long
            For ClassNo = .LowerClass To .UpperClass
                Dim ClassTotal As Long
                ClassTotal = Member.SumMainStreamClass(Grade, ClassNo)
                .ClassTotal(Grade, ClassNo) = ClassTotal
            Next
        Next
    End With
End Sub

Private Sub AggregateSpecial(ByVal Member As Dom_EnrollMember)
    With Special
        Dim Grade As Long
        For Grade = .LowerGrade To .UpperGrade
            Dim MaleGradeTotal As Long
            MaleGradeTotal = Member.SumSpecialMale(Grade)
            Dim FemaleGradeTotal As Long
            FemaleGradeTotal = Member.SumSpecialFemale(Grade)
            .MaleGradeTotal(Grade) = MaleGradeTotal
            .FemaleGradeTotal(Grade) = FemaleGradeTotal
            .MaleGrandTotal = .MaleGrandTotal + MaleGradeTotal
            .FemaleGrandTotal = .FemaleGrandTotal + FemaleGradeTotal
            Dim ClassNo As Long
            For ClassNo = .LowerClass To .UpperClass
                Dim ClassTotal As Long
                ClassTotal = Member.SumSpecialClass(Grade, ClassNo)
                .ClassTotal(Grade, ClassNo) = ClassTotal
            Next
        Next
    End With
End Sub

