VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TransferMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal Transfer As Dom_Transfer) As String
    ToRaw = CStr(Transfer.IsTrue)
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Transfer
    Set ToDomain = This.Factory.CreateTransfer(IsTrue(Raw))
End Function

Public Function ToModel(ByVal Raw As String) As String
    ToModel = IsTrue(Raw)
End Function

Private Function IsTrue(ByVal Raw As String) As Boolean
    IsTrue = Inf_BooleanUtility.ToBoolean(Raw)
End Function
