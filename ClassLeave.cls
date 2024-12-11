VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("FormExBB")
Option Explicit
Private DataBody_ As Variant
Private Const KEY_取得者    As String = "Person"
Private Const KEY_種別      As String = "LeaveType"
Private Const KEY_開始日    As String = "FirstDate"
Private Const KEY_終了日    As String = "LastDate"
Private Const KEY_表示      As String = "View"
Private Const TEXT_取得者   As String = "取得者"
Private Const TEXT_種別     As String = "種別"
Private Const TEXT_開始日   As String = "開始日"
Private Const TEXT_終了日   As String = "終了日"
Private Const TEXT_表示     As String = "表示"
Private Enum Header
    Columns取得者 = 1
    Columns種別
    Columns開始日
    Columns終了日
    Columns表示
End Enum
Private Sub Class_Initialize()
    Dim FilePath As String
    FilePath = CSO.GetPath.data.leave.FileLeave
    If Not FSO.FileExists(FilePath) Then Exit Sub
    Me.DataBody = LIB.ReadAllLineAsArray2D(FilePath)
End Sub
Public Property Get DataBody() As Variant
    DataBody = DataBody_
End Property
Public Property Let DataBody(ByVal Target As Variant)
    DataBody_ = Target
End Property
Public Function Item取得者(ByVal R As Long) As String
    Item取得者 = DataBody(R, Me.Index取得者)
End Function
Public Function Item種別(ByVal R As Long) As String
    Item種別 = DataBody(R, Me.Index種別)
End Function
Public Function Item開始日(ByVal R As Long) As Date
    Item開始日 = DataBody(R, Me.Index開始日)
End Function
Public Function Item終了日(ByVal R As Long) As Date
    Item終了日 = DataBody(R, Me.Index終了日)
End Function
Public Function Item表示(ByVal R As Long) As String
    Item表示 = DataBody(R, Me.Index表示)
End Function
Public Function Index取得者() As Long
    Index取得者 = Header.Columns取得者
End Function
Public Function Index種別() As Long
    Index種別 = Header.Columns種別
End Function
Public Function Index開始日() As Long
    Index開始日 = Header.Columns開始日
End Function
Public Function Index終了日() As Long
    Index終了日 = Header.Columns終了日
End Function
Public Function Index表示() As Long
    Index表示 = Header.Columns表示
End Function
Public Function Key取得者() As String
    Key取得者 = KEY_取得者
End Function
Public Function Key種別() As String
    Key種別 = KEY_種別
End Function
Public Function Key開始日() As String
    Key開始日 = KEY_開始日
End Function
Public Function Key終了日() As String
    Key終了日 = KEY_終了日
End Function
Public Function Key表示() As String
    Key表示 = KEY_表示
End Function
Public Function Text取得者() As String
    Text取得者 = TEXT_取得者
End Function
Public Function Text種別() As String
    Text種別 = TEXT_種別
End Function
Public Function Text開始日() As String
    Text開始日 = TEXT_開始日
End Function
Public Function Text終了日() As String
    Text終了日 = TEXT_終了日
End Function
Public Function Text表示() As String
    Text表示 = TEXT_表示
End Function
Public Function IsValidDate(ByVal BasisDate As Date, ByVal FirstDate As Date, ByVal LastDate As Date) As Boolean
    Dim Apr1st As Date
    Dim Mar31th As Date
    Apr1st = LIB.Date4月1日(BasisDate)
    Mar31th = LIB.Date3月31日(BasisDate)
    If Apr1st <= FirstDate And FirstDate <= Mar31th Then
        If Apr1st <= LastDate And LastDate <= Mar31th Then
            If FirstDate <= LastDate Then IsValidDate = True
        End If
    End If
End Function
