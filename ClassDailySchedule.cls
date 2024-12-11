VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDailySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Schdule")
Option Explicit
Private BasisDate_      As Date
Private RecordNumbers_  As Collection
Private Headers_        As Variant
Private ItemNames_      As Dictionary
Private DataBody_       As Variant
Public Property Let BasisDate(ByVal Target As Date)
    BasisDate_ = Target
End Property
Public Property Get BasisDate() As Date
    BasisDate = BasisDate_
End Property
Public Function RecordNumbers() As Collection
    Set RecordNumbers = RecordNumbers_
End Function
Public Sub GetReadyRecordNumbers()
    Set RecordNumbers_ = AddRecordNumbers
End Sub
Public Property Let Headers(ByVal Target As Variant)
    Headers_ = Target
End Property
Public Property Get Headers() As Variant
    Headers = Headers_
End Property
Public Property Get ItemNames() As Dictionary
    Set ItemNames = ItemNames_
End Property
Public Property Set ItemNames(ByVal Target As Dictionary)
    Set ItemNames_ = Target
End Property
Public Property Get DataBody() As Variant
    DataBody = DataBody_
End Property
Public Property Let DataBody(ByVal Target As Variant)
    DataBody_ = Target
End Property
Public Function RecordNumber() As Long
    RecordNumber = LIB.DateDiff4åé1ì˙(Me.BasisDate)
End Function
Public Function RecordDataBody() As Variant()
    Dim Target As Range
    Set Target = Me.RecordRange
    If Target Is Nothing Then Exit Function
    RecordDataBody = Target.Value
End Function
Public Function RecordRange() As Range
    If Schedule.TargetTable Is Nothing Then Exit Function
    Set RecordRange = Schedule.TargetTable.ListRows.Item(Me.RecordNumber).Range
End Function
Public Function ClassWeeks() As String
    ClassWeeks = Schedule.CountedClassWeeks(Me.BasisDate)
End Function
Public Function ThisWeekNoticeRecordNumber() As Long
    ThisWeekNoticeRecordNumber = LIB.DateDiff4åé1ì˙(ThisWeekNoticeDate)
End Function
Public Function NextWeekNoticeRecordNumber() As Long
    NextWeekNoticeRecordNumber = LIB.DateDiff4åé1ì˙(NextWeekNoticeDate)
End Function
Private Function AddRecordNumbers() As Collection
    Dim RET As Collection
    Set RET = New Collection
    RET.Add RecordNumber
    Set AddRecordNumbers = RET
End Function
Private Function ThisWeekNoticeDate() As Date
    Dim RET As Date
    If LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday) < LIB.Date4åé1ì˙(Me.BasisDate) Then
        RET = LIB.Date4åé1ì˙(Me.BasisDate)
    Else
        RET = LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday)
    End If
    ThisWeekNoticeDate = RET
End Function
Private Function NextWeekNoticeDate() As Date
    Dim RET As Date
    If LIB.Date3åé31ì˙(Me.BasisDate) < LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday) + 7 Then
        RET = LIB.Date3åé31ì˙(Me.BasisDate)
    Else
        RET = LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday) + 7
    End If
    NextWeekNoticeDate = RET
End Function
