VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entry")
Option Explicit

Private Type Member
    DateID  As Dom_DateID
    Name    As Dom_ScheduleName
    Value   As Dom_ScheduleValue
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal Name As Dom_ScheduleName, ByVal Value As Dom_ScheduleValue)
    Set This.DateID = DateID
    Set This.Name = Name
    Set This.Value = Value
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_ScheduleValue)
    Set This.Value = NewValue
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Property Get Name() As Dom_ScheduleName
    Set Name = This.Name
End Property

Public Property Get Value() As Dom_ScheduleValue
    Set Value = This.Value
End Property
