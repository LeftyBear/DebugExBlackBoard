VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TransferMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Public Function ToRaw(ByVal Transfer As Dom_Transfer) As String
    ToRaw = CStr(Transfer.IsTrue)
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Transfer
    Set ToDomain = Dom_ValueObjectFactory.CreateTransfer(GetState(Raw))
End Function

Public Function ToModel(ByVal Raw As String) As String
    ToModel = GetState(Raw)
End Function

Private Function GetState(ByVal Raw As String) As Boolean
    GetState = Inf_BooleanUtility.ToBoolean(Raw)
End Function
