VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentHeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Parser")
Option Explicit
Implements Dom_IEnrollmentColumnParser

Private Function Dom_IEnrollmentColumnParser_Parse(ByVal HeaderName As String) As App_EnrollmentColumnDTO
    If HeaderName = "日付" Then
        Set Dom_IEnrollmentColumnParser_Parse = CreateSpecialDTO(1)
        Exit Function
    End If
    If HeaderName = "在籍異動" Then
        Set Dom_IEnrollmentColumnParser_Parse = CreateSpecialDTO(2)
        Exit Function
    End If
    If HeaderName = "在籍備考" Then
        Set Dom_IEnrollmentColumnParser_Parse = CreateSpecialDTO(3)
        Exit Function
    End If
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, Util_Character.UnderScore)
    If UBound(Parts) < 2 Then Err.Raise App_ErrNum.InvalidHeaderName, , "Header name is invalid"
    Dim DTO As App_EnrollmentColumnDTO
    Set DTO = New App_EnrollmentColumnDTO
    DTO.Initialize ExtractNumber(Parts(0)), ExtractNumber(Parts(1)), ConvertGender(Parts(2)), 0, False
    Set Dom_IEnrollmentColumnParser_Parse = DTO
End Function

Private Function CreateSpecialDTO(ByVal Code As Long) As App_EnrollmentColumnDTO
    Dim DTO As App_EnrollmentColumnDTO
    Set DTO = New App_EnrollmentColumnDTO
    DTO.Initialize 0, 0, 0, Code, True
    Set CreateSpecialDTO = DTO
End Function

Private Function ExtractNumber(ByVal Part As String) As Long
    ExtractNumber = CLng(VBA.Replace(VBA.Replace(Part, "年", vbNullString), "組", vbNullString))
End Function

Private Function ConvertGender(ByVal Part As String) As Long
    Select Case Part
    Case "男":  ConvertGender = 1
    Case "女":  ConvertGender = 2
    End Select
End Function
