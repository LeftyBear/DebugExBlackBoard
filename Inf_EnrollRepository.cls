VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit

Private Type Member
    Reader As Inf_IReadRepository
    Config As App_EnrollConfig
End Type
Private This As Member

Public Sub Initialize(ByVal Reader As Inf_IReadRepository, ByVal Config As App_EnrollConfig)
    Set This.Reader = Reader
    Set This.Config = Config
End Sub

Public Function ReadAll() As VBA.Collection
    Dim Raw As VBA.Collection
    Set Raw = This.Reader.ReadAll
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Row As Scripting.Dictionary
    For Each Row In Raw
        Dim Member As Dom_EnrollMember
        Set Member = MapRow(Row)
        Result.Add Member
    Next
    Set ReadAll = Result
End Function

Private Function MapRow(ByVal Row As Scripting.Dictionary) As Dom_EnrollMember
    Dim Member As Dom_EnrollMember
    Set Member = New Dom_EnrollMember
    Dim MainStreamMale() As Long
    Dim MainStreamFemale() As Long
    MapMainStream Row, MainStreamMale, MainStreamFemale
    Dim SpecialMale() As Long
    Dim SpecialFemale() As Long
    MapSpecial Row, SpecialMale, SpecialFemale
    Dim TargetDate As Date
    TargetDate = Util_Function.NormalizeToDate(Row.Item(This.Config.TargetDate))
    Dim Transfer As Boolean
    Transfer = Util_Function.NormalizeToBoolean(Row.Item(This.Config.Transfer))
    Dim Remarks As String
    Remarks = Row.Item(This.Config.Remarks)
    Member.Inject TargetDate, MainStreamMale, MainStreamFemale, SpecialMale, SpecialFemale, Transfer, Remarks
    Set MapRow = Member
End Function

Private Sub MapMainStream(ByVal Row As Scripting.Dictionary, ByRef MainStreamMale() As Long, ByRef MainStreamFemale() As Long)
    ReDim MainStreamMale(1 To This.Config.MainStreamMaxGrade, 1 To This.Config.MainStreamMaxClass)
    ReDim MainStreamFemale(1 To This.Config.MainStreamMaxGrade, 1 To This.Config.MainStreamMaxClass)
    Dim Grade As Long
    For Grade = 1 To This.Config.MainStreamMaxGrade
        Dim ClassNo As Long
        For ClassNo = 1 To This.Config.MainStreamMaxClass
            Dim Buffer As Long
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.MainStreamMaleColumnName(Grade, ClassNo)))
            MainStreamMale(Grade, ClassNo) = CLng(Buffer)
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.MainStreamFemaleColumnName(Grade, ClassNo)))
            MainStreamFemale(Grade, ClassNo) = CLng(Buffer)
        Next
    Next
End Sub

Private Sub MapSpecial(ByVal Row As Scripting.Dictionary, ByRef SpecialMale() As Long, ByRef SpecialFemale() As Long)
    ReDim SpecialMale(1 To This.Config.SpecialMaxGrade, 1 To This.Config.SpecialMaxClass)
    ReDim SpecialFemale(1 To This.Config.SpecialMaxGrade, 1 To This.Config.SpecialMaxClass)
    Dim Grade As Long
    For Grade = 1 To This.Config.SpecialMaxGrade
        Dim ClassNo As Long
        For ClassNo = 1 To This.Config.SpecialMaxClass
            Dim Buffer As Long
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.SpecialMaleColumnName(Grade, ClassNo)))
            SpecialMale(Grade, ClassNo) = CLng(Buffer)
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.SpecialFemaleColumnName(Grade, ClassNo)))
            SpecialFemale(Grade, ClassNo) = CLng(Buffer)
        Next
    Next
End Sub
