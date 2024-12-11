VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormSizingItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private FormLeft_       As Single
Private FormTop_        As Single
Private FormWidth_      As Single
Private FormHeight_     As Single
Private RateFontSize_   As Single
Public Property Get FormResizingItem() As Object
    Set FormResizingItem = Me
End Property
Public Property Get FormLeft() As Single
    FormLeft = FormLeft_
End Property
Public Property Let FormLeft(ByVal Figures As Single)
    FormLeft_ = Figures
End Property
Public Property Get FormTop() As Single
    FormTop = FormTop_
End Property
Public Property Let FormTop(ByVal Figures As Single)
    FormTop_ = Figures
End Property
Public Property Get FormWidth() As Single
    FormWidth = FormWidth_
End Property
Public Property Let FormWidth(ByVal Figures As Single)
    FormWidth_ = Figures
End Property
Public Property Get FormHeight() As Single
    FormHeight = FormHeight_
End Property
Public Property Let FormHeight(ByVal Figures As Single)
    FormHeight_ = Figures
End Property
Public Property Get RateFontSize() As Single
    RateFontSize = RateFontSize_
End Property
Public Property Let RateFontSize(ByVal Figures As Single)
    RateFontSize_ = Figures
End Property

