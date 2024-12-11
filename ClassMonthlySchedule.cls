VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMonthlySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Schdule")
Option Explicit
Private BasisDate_      As Date
Private RecordNumbers_  As Collection
Private ItemNames_      As Dictionary
Private DataBody_       As Variant
Private Const MAX_DAYS  As Long = 31
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
Private Function AddRecordNumbers() As Collection
    Dim RET As Collection
    Set RET = New Collection
    Dim RecordNumber As Long
    RecordNumber = FirstRecordNumber
    Dim Days As Long
    For Days = 1 To MAX_DAYS
        If RecordNumber <= LastRecordNumber Then
            RET.Add RecordNumber
            RecordNumber = RecordNumber + 1
        Else
            RET.Add 0
        End If
    Next
    Set AddRecordNumbers = RET
End Function
Private Function FirstRecordNumber() As Long
    FirstRecordNumber = LIB.DateDiff4ŒŽ1“ú(FirstDate)
End Function
Private Function LastRecordNumber() As Long
    LastRecordNumber = LIB.DateDiff4ŒŽ1“ú(LastDate)
End Function
Private Function FirstDate() As Date
    FirstDate = LIB.Date¡ŒŽ1“ú(Me.BasisDate)
End Function
Private Function LastDate() As Date
    LastDate = LIB.Date¡ŒŽ––“ú(Me.BasisDate)
End Function
