VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Enrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.Entity"
Option Explicit

Private Type MainStreamMember
    Male()      As Long
    Female()    As Long
End Type

Private Type SpecialMember
    Male()      As Long
    Female()    As Long
End Type

Private Type Member
    TargetDate  As Date
    MainStream  As MainStreamMember
    Special     As SpecialMember
    Transfer    As Boolean
    Remarks     As String
End Type

Private This As Member

Friend Sub Inject(ByVal TargetDate As Date, ByRef MainStreamMale() As Long, ByRef MainStreamFemale() As Long, ByRef SpecialMale() As Long, ByRef SpecialFemale() As Long, ByVal Transfer As Boolean, ByVal Remarks As String)
    This.TargetDate = TargetDate
    ReDim This.MainStream.Male(LBound(MainStreamMale, D1) To UBound(MainStreamMale, D1), LBound(MainStreamMale, D2) To UBound(MainStreamMale, D2))
    ReDim This.MainStream.Female(LBound(MainStreamFemale, D1) To UBound(MainStreamFemale, D1), LBound(MainStreamFemale, D2) To UBound(MainStreamFemale, D2))
    ReDim This.Special.Male(LBound(SpecialMale, D1) To UBound(SpecialMale, D1), LBound(SpecialMale, D2) To UBound(SpecialMale, D2))
    ReDim This.Special.Female(LBound(SpecialFemale, D1) To UBound(SpecialFemale, D1), LBound(SpecialFemale, D2) To UBound(SpecialFemale, D2))
    This.MainStream.Male = CopyLongArray(MainStreamMale)
    This.MainStream.Female = CopyLongArray(MainStreamFemale)
    This.Special.Male = CopyLongArray(SpecialMale)
    This.Special.Female = CopyLongArray(SpecialFemale)
    This.Transfer = Transfer
    This.Remarks = Remarks
End Sub

Public Property Get TargetDate() As Date
    TargetDate = This.TargetDate
End Property

Public Property Get MainStreamUpperGrade() As Long
    MainStreamUpperGrade = UBound(This.MainStream.Male, D1)
End Property

Public Property Get MainStreamLowerGrade() As Long
    MainStreamLowerGrade = LBound(This.MainStream.Male, D1)
End Property

Public Property Get MainStreamUpperClass() As Long
    MainStreamUpperClass = UBound(This.MainStream.Male, D2)
End Property

Public Property Get MainStreamLowerClass() As Long
    MainStreamLowerClass = LBound(This.MainStream.Male, D2)
End Property

Public Property Get SpecialUpperGrade() As Long
    SpecialUpperGrade = UBound(This.Special.Male, D1)
End Property

Public Property Get SpecialLowerGrade() As Long
    SpecialLowerGrade = LBound(This.Special.Male, D1)
End Property

Public Property Get SpecialUpperClass() As Long
    SpecialUpperClass = UBound(This.Special.Male, D2)
End Property

Public Property Get SpecialLowerClass() As Long
    SpecialLowerClass = LBound(This.Special.Male, D2)
End Property

Public Function MainStreamMale(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateMainStreamGrade Grade
    ValidateMainStreamClass ClassNo
    MainStreamMale = This.MainStream.Male(Grade, ClassNo)
End Function

Public Function MainStreamFemale(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateMainStreamGrade Grade
    ValidateMainStreamClass ClassNo
    MainStreamFemale = This.MainStream.Female(Grade, ClassNo)
End Function

Public Function SumMainStreamMale(ByVal Grade As Long) As Long
    ValidateMainStreamGrade Grade
    SumMainStreamMale = SumInnerGender(This.MainStream.Male, Grade)
End Function

Public Function SumMainStreamFemale(ByVal Grade As Long) As Long
    ValidateMainStreamGrade Grade
    SumMainStreamFemale = SumInnerGender(This.MainStream.Female, Grade)
End Function

Public Function SumMainStreamGrade(ByVal Grade As Long) As Long
    ValidateMainStreamGrade Grade
    SumMainStreamGrade = Me.SumMainStreamMale(Grade) + Me.SumMainStreamFemale(Grade)
End Function

Public Function SumMainStreamClass(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateMainStreamGrade Grade
    ValidateMainStreamClass ClassNo
    SumMainStreamClass = This.MainStream.Male(Grade, ClassNo) + This.MainStream.Female(Grade, ClassNo)
End Function

Public Function SpecialMale(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateSpecialGrade Grade
    ValidateSpecialClass ClassNo
    SpecialMale = This.Special.Male(Grade, ClassNo)
End Function

Public Function SpecialFemale(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateSpecialGrade Grade
    ValidateSpecialClass ClassNo
    SpecialFemale = This.Special.Female(Grade, ClassNo)
End Function

Public Function SumSpecialMale(ByVal Grade As Long) As Long
    ValidateSpecialGrade Grade
    SumSpecialMale = SumInnerGender(This.Special.Male, Grade)
End Function

Public Function SumSpecialFemale(ByVal Grade As Long) As Long
    ValidateSpecialGrade Grade
    SumSpecialFemale = SumInnerGender(This.Special.Female, Grade)
End Function

Public Function SumSpecialGrade(ByVal Grade As Long) As Long
    ValidateSpecialGrade Grade
    SumSpecialGrade = Me.SumSpecialMale(Grade) + Me.SumSpecialFemale(Grade)
End Function

Public Function SumSpecialClass(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    ValidateSpecialGrade Grade
    ValidateSpecialClass ClassNo
    SumSpecialClass = This.Special.Male(Grade, ClassNo) + This.Special.Female(Grade, ClassNo)
End Function

Private Function SumInnerGender(ByRef Values() As Long, ByVal Grade As Long) As Long
    Dim ClassNo As Long
    For ClassNo = LBound(Values, D2) To UBound(Values, D2)
        Dim Result As Long
        Result = Result + Values(Grade, ClassNo)
    Next
    SumInnerGender = Result
End Function

Public Function Transfer() As Boolean
    Transfer = This.Transfer
End Function

Public Function Remarks() As String
    Remarks = This.Remarks
End Function

Private Sub ValidateMainStreamGrade(ByVal Grade As Long)
    If Grade < LBound(This.MainStream.Male, D1) Or UBound(This.MainStream.Male, D1) < Grade Then
        Err.Raise Dom_ErrNum.OutOfRange, , "Grade is out of range"
    End If
End Sub

Private Sub ValidateMainStreamClass(ByVal ClassNo As Long)
    If ClassNo < LBound(This.MainStream.Male, D2) Or UBound(This.MainStream.Male, D2) < ClassNo Then
        Err.Raise Dom_ErrNum.OutOfRange, , "ClassNo is out of range"
    End If
End Sub

Private Sub ValidateSpecialGrade(ByVal Grade As Long)
    If Grade < LBound(This.Special.Male, D1) Or UBound(This.Special.Male, D1) < Grade Then
        Err.Raise Dom_ErrNum.OutOfRange, , "Grade is out of range"
    End If
End Sub

Private Sub ValidateSpecialClass(ByVal ClassNo As Long)
    If ClassNo < LBound(This.Special.Male, D2) Or UBound(This.Special.Male, D2) < ClassNo Then
        Err.Raise Dom_ErrNum.OutOfRange, , "ClassNo is out of range"
    End If
End Sub

Private Function CopyLongArray(ByRef Values() As Long) As Long()
    Dim Result() As Long
    ReDim Result(LBound(Values, D1) To UBound(Values, D1), LBound(Values, D2) To UBound(Values, D2))
    Dim R As Long
    For R = LBound(Values, D1) To UBound(Values, D1)
        Dim C As Long
        For C = LBound(Values, D2) To UBound(Values, D2)
            Result(R, C) = Values(R, C)
        Next
    Next
    CopyLongArray = Result
End Function

