VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWordSearching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Schdule")
Option Explicit
Private FoundItemNameList_ As Dictionary
Public Property Set FoundItemNameList(ByVal Target As Dictionary)
    Set FoundItemNameList_ = Target
End Property
Public Property Get FoundItemNameList() As Dictionary
    Set FoundItemNameList = FoundItemNameList_
End Property

