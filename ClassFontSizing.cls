VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFontSizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal Hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal Hwnd As LongPtr, ByVal hDc As LongPtr) As Long
Private Sheet_              As Worksheet
Private WidthPerLen1Point_  As Double
Private WidthPerLen2Point_  As Double
Public Property Set Sheet(ByVal Target As Worksheet)
    Set Sheet_ = Target
End Property
Public Property Get Sheet() As Worksheet
    Set Sheet = Sheet_
End Property
Private Property Let WidthPerLen1Point(ByVal Point As Double)
    WidthPerLen1Point_ = Point
End Property
Private Property Get WidthPerLen1Point() As Double
    WidthPerLen1Point = WidthPerLen1Point_
End Property
Private Property Let WidthPerLen2Point(ByVal Point As Double)
    WidthPerLen2Point_ = Point
End Property
Private Property Get WidthPerLen2Point() As Double
    WidthPerLen2Point = WidthPerLen2Point_
End Property
Public Function DownedFontSize(ByVal Text As String, ByVal SingleCell As Range, ByVal MaxFontSize As Double, ByVal MinFontSize As Double) As Double
    Dim FontSizePoint As Double
    For FontSizePoint = MaxFontSize To MinFontSize Step -0.5
        If Not Me.ShouldDown(Text, SingleCell, FontSizePoint) Then Exit For
    Next
    DownedFontSize = FontSizePoint
End Function
Public Function ShouldDown(ByVal Text As String, ByVal SingleCell As Range, Optional ByVal FontSizePoint As Double) As Boolean
    If Text Like vbNullString Then Exit Function
    Dim Texts() As String
    Texts = Split(Text, vbLf)
    Dim ColumnLength As Long
    ColumnLength = CalculateColumnLength(SingleCell, FontSizePoint)
    Dim i As Long
    For i = LBound(Texts) To UBound(Texts)
        Dim LineCount As Long
        LineCount = LineCount + CalculateLinesRequiredInDisplay(Texts(i), ColumnLength)
    Next
    If LineCount <= CalculateRowLines(SingleCell, FontSizePoint) Then Exit Function
    ShouldDown = True
End Function
Private Function CalculateColumnLength(ByVal SingleCell As Range, Optional ByVal FontSizePoint As Double) As Long
    Dim RET As Long
    If FontSizePoint > 0 Then
        RET = Int((RangeWidthPixel(SingleCell)) / (Length1WidthPixel * FontSizeRate(FontSizePoint)))
    Else
        RET = Int((RangeWidthPixel(SingleCell)) / (Length1WidthPixel * FontSizeRate(SingleCell.Font.Size)))
    End If
    CalculateColumnLength = RET
End Function
Private Function RangeWidthPixel(ByVal SingleCell As Range) As Long
    RangeWidthPixel = ToPixel(SingleCell.MergeArea.Width) - WhitespacePixel
End Function
Private Function Length1WidthPixel() As Long
    Length1WidthPixel = Length2WidthPixelWithWhitespace - Length1WidthPixelWithWhitespace
End Function
Private Function WhitespacePixel() As Long
    WhitespacePixel = Length1WidthPixelWithWhitespace - Length1WidthPixel
End Function
Private Function Length1WidthPixelWithWhitespace() As Long
    If WidthPerLen1Point = 0 Then
        With Me.Sheet.Columns.Item(100)
            Dim TemporaryColumnWidth As Double
            TemporaryColumnWidth = .ColumnWidth
            .ColumnWidth = 1
            WidthPerLen1Point = .Width
            .ColumnWidth = TemporaryColumnWidth
        End With
    End If
    Length1WidthPixelWithWhitespace = ToPixel(WidthPerLen1Point)
End Function
Private Function Length2WidthPixelWithWhitespace() As Long
    If WidthPerLen2Point = 0 Then
        With Me.Sheet.Columns.Item(100)
            Dim TemporaryColumnWidth As Double
            TemporaryColumnWidth = .ColumnWidth
            .ColumnWidth = 2
            WidthPerLen2Point = .Width
            .ColumnWidth = TemporaryColumnWidth
        End With
    End If
    Length2WidthPixelWithWhitespace = ToPixel(WidthPerLen2Point)
End Function
Private Function CalculateLinesRequiredInDisplay(ByVal Text As String, ByVal ColumnLength As Long) As Long
    CalculateLinesRequiredInDisplay = WorksheetFunction.RoundUp(LenB(StrConv(Text, vbFromUnicode)) / ColumnLength, 0)
End Function
Private Function CalculateRowLines(ByVal SingleCell As Range, Optional ByVal FontSizePoint As Double) As Long
    Dim RET As Long
    If FontSizePoint > 0 Then
        RET = Int(RangeHeightPixel(SingleCell) / (Length1HeightPixel * FontSizeRate(FontSizePoint)))
    Else
        RET = Int(RangeHeightPixel(SingleCell) / (Length1HeightPixel * FontSizeRate(SingleCell.Font.Size)))
    End If
    CalculateRowLines = RET
End Function
Private Function RangeHeightPixel(ByVal SingleCell As Range) As Long
    RangeHeightPixel = ToPixel(SingleCell.MergeArea.Height)
End Function
Private Function Length1HeightPixel() As Long
    Length1HeightPixel = ToPixel(Me.Sheet.StandardHeight)
End Function
Private Function FontSizeRate(ByVal FontSizePoint As Double) As Double
    FontSizeRate = FontSizePoint / Application.StandardFontSize
End Function
Private Function ToPixel(ByVal Point As Double) As Long
    ToPixel = CLng(Point / ConvertingRate)
End Function
Private Function ConvertingRate() As Double
    Static StaticRate As Double
    If StaticRate = 0 Then
        Dim WindowHandle As LongPtr
        Dim DCHandle As LongPtr
        WindowHandle = Excel.Application.Hwnd
        DCHandle = GetDC(WindowHandle)
        StaticRate = 72 / GetDeviceCaps(DCHandle, &H58&)
        ReleaseDC WindowHandle, DCHandle
    End If
    ConvertingRate = StaticRate
End Function

