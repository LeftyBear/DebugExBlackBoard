VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPrintingItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Schdule")
Option Explicit
Private HeaderBody_                 As Variant
Private DataBody_                   As Variant
Private Const KEY_項目名            As String = "ItemName"
Private Const KEY_印刷範囲          As String = "PrintingArea"
Private Const KEY_印刷月            As String = "PrintingMonth"
Private Const KEY_印刷週            As String = "PrintingWeek"
Private Const KEY_集計方法          As String = "TotalType"
Private Const KEY_文字サイズ縮小    As String = "ShouldDownFontSize"
Private Const KEY_項目名追加        As String = "ShouldAddItemName"
Private Const TEXT_項目名           As String = "項目名"
Private Const TEXT_印刷範囲         As String = "印刷範囲"
Private Const TEXT_印刷月           As String = "印刷月"
Private Const TEXT_印刷週           As String = "印刷週"
Private Const TEXT_集計方法         As String = "集計方法"
Private Const TEXT_文字サイズ縮小   As String = "文字サイズ縮小"
Private Const TEXT_項目名追加       As String = "項目名追加"
Private Enum Header
    '@Ignore UseMeaningfulName
    項目名 = 1
    '@Ignore UseMeaningfulName
    印刷範囲
    '@Ignore UseMeaningfulName
    印刷月
    '@Ignore UseMeaningfulName
    印刷週
    '@Ignore UseMeaningfulName
    集計方法
    '@Ignore UseMeaningfulName
    文字サイズ縮小
    '@Ignore UseMeaningfulName
    項目名追加
    MaxColumns
End Enum
Public Sub ReadFile(ByVal FilePath As String)
    If Not FSO.FileExists(FilePath) Then Exit Sub
    Dim TMP() As Variant
    TMP = LIB.ReadAllLineAsArray2D(FilePath)
    If LIB.NotIsArray(TMP) Then Exit Sub
    Dim BUF() As Variant
    ReDim BUF(1 To 1, 1 To UBound(TMP, 2))
    Dim R As Long
    Dim C As Long
    R = 1
    For C = 1 To UBound(BUF, 2)
        BUF(R, C) = TMP(R, C)
    Next
    Me.HeaderBody = BUF
    ReDim BUF(1 To UBound(TMP, 1) - 1, 1 To UBound(TMP, 2))
    Dim SpecialCase As Long
    SpecialCase = Me.C印刷範囲
    For R = 1 To UBound(BUF, 1)
        For C = 1 To UBound(BUF, 2)
            If C = SpecialCase Then
                BUF(R, C) = Replace(TMP(R + 1, C), SEMICOLON, CANNMA)
            Else
                BUF(R, C) = TMP(R + 1, C)
            End If
        Next
    Next
    Me.DataBody = BUF
End Sub
Public Property Let HeaderBody(ByVal Target As Variant)
    HeaderBody_ = Target
End Property
'TODO ヘッダーをリストビューに反映させる
'@Ignore ProcedureNotUsed
Public Property Get HeaderBody() As Variant
    HeaderBody = HeaderBody_
End Property
Public Property Let DataBody(ByVal Target As Variant)
    DataBody_ = Target
End Property
Public Property Get DataBody() As Variant
    DataBody = DataBody_
End Property
Public Function Item項目名(ByVal R As Long) As String
    If Me.C項目名 > ColumnsCount Then Exit Function
    Item項目名 = Me.DataBody(R, Me.C項目名)
End Function
Public Function Item印刷範囲(ByVal R As Long) As String
    If Me.C印刷範囲 > ColumnsCount Then Exit Function
    Item印刷範囲 = Replace(Me.DataBody(R, Me.C印刷範囲), SEMICOLON, CANNMA)
End Function
Public Function Item印刷月(ByVal R As Long) As String
    If Me.C印刷月 > ColumnsCount Then Exit Function
    Item印刷月 = Me.DataBody(R, Me.C印刷月)
End Function
Public Function Item印刷週(ByVal R As Long) As String
    If Me.C印刷週 > ColumnsCount Then Exit Function
    Item印刷週 = Me.DataBody(R, Me.C印刷週)
End Function
Public Function Item集計方法(ByVal R As Long) As String
    If Me.C集計方法 > ColumnsCount Then Exit Function
    Item集計方法 = Me.DataBody(R, Me.C集計方法)
End Function
Public Function Item文字サイズ縮小(ByVal R As Long) As String
    If Me.C文字サイズ縮小 > ColumnsCount Then Exit Function
    Item文字サイズ縮小 = Me.DataBody(R, Me.C文字サイズ縮小)
End Function
Public Function Item項目名追加(ByVal R As Long) As String
    If Me.C項目名追加 > ColumnsCount Then Exit Function
    Item項目名追加 = Me.DataBody(R, Me.C項目名追加)
End Function
Public Function Key項目名() As String
    Key項目名 = KEY_項目名
End Function
Public Function Key印刷範囲() As String
    Key印刷範囲 = KEY_印刷範囲
End Function
Public Function Key印刷月() As String
    Key印刷月 = KEY_印刷月
End Function
Public Function Key印刷週() As String
    Key印刷週 = KEY_印刷週
End Function
Public Function Key集計方法() As String
    Key集計方法 = KEY_集計方法
End Function
Public Function Key文字サイズ縮小() As String
    Key文字サイズ縮小 = KEY_文字サイズ縮小
End Function
Public Function Key項目名追加() As String
    Key項目名追加 = KEY_項目名追加
End Function
Public Function Text項目名() As String
    Text項目名 = TEXT_項目名
End Function
Public Function Text印刷範囲() As String
    Text印刷範囲 = TEXT_印刷範囲
End Function
Public Function Text印刷月() As String
    Text印刷月 = TEXT_印刷月
End Function
Public Function Text印刷週() As String
    Text印刷週 = TEXT_印刷週
End Function
Public Function Text集計方法() As String
    Text集計方法 = TEXT_集計方法
End Function
Public Function Text文字サイズ縮小() As String
    Text文字サイズ縮小 = TEXT_文字サイズ縮小
End Function
Public Function Text項目名追加() As String
    Text項目名追加 = TEXT_項目名追加
End Function
'@Ignore UseMeaningfulName
Public Function C項目名() As Long
    C項目名 = Header.項目名
End Function
'@Ignore UseMeaningfulName
Public Function C印刷範囲() As Long
    C印刷範囲 = Header.印刷範囲
End Function
'@Ignore UseMeaningfulName
Public Function C印刷月() As Long
    C印刷月 = Header.印刷月
End Function
'@Ignore UseMeaningfulName
Public Function C印刷週() As Long
    C印刷週 = Header.印刷週
End Function
'@Ignore UseMeaningfulName
Public Function C集計方法() As Long
    C集計方法 = Header.集計方法
End Function
'@Ignore UseMeaningfulName
Public Function C文字サイズ縮小() As Long
    C文字サイズ縮小 = Header.文字サイズ縮小
End Function
'@Ignore UseMeaningfulName
Public Function C項目名追加() As Long
    C項目名追加 = Header.項目名追加
End Function
Public Function ColumnsCount() As Long
    ColumnsCount = Header.MaxColumns - 1
End Function
