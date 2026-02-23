VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentCompositeColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit
Implements Dom_IEnrollmentColumn
Private Type Member
    StreamType      As Dom_StreamType
    GradeColumn     As Dom_GradeColumn
    ClassColumn     As Dom_ClassColumn
    GenderColumn    As Dom_GenderColumn
    Value           As String
End Type

Private This As Member

Public Sub Initialize(ByVal StreamType As Dom_StreamType, ByVal GradeColumn As Dom_GradeColumn, ByVal ClassColumn As Dom_ClassColumn, ByVal GenderColumn As Dom_GenderColumn, ByVal Value As String)
    If GradeColumn Is Nothing Then Err.Raise errNothingGradeColumn, , "GradeColumn is Nothing"
    If ClassColumn Is Nothing Then Err.Raise errNothingClassColumn, , "ClassColumn is Nothing"
    If GenderColumn Is Nothing Then Err.Raise errNothingGenderColumn, , "GenderColumn is Nothing"
    This.StreamType = StreamType
    Set This.GradeColumn = GradeColumn
    Set This.ClassColumn = ClassColumn
    Set This.GenderColumn = GenderColumn
    This.Value = Value
End Sub

Public Property Get StreamType() As Dom_StreamType
    StreamType = This.StreamType
End Property

Public Property Get GradeColumn() As Dom_GradeColumn
    Set GradeColumn = This.GradeColumn
End Property

Public Property Get ClassColumn() As Dom_ClassColumn
    Set ClassColumn = This.ClassColumn
End Property

Public Property Get GenderColumn() As Dom_GenderColumn
    Set GenderColumn = This.GenderColumn
End Property

Public Property Get Value() As String
    Value = This.Value
End Property

Private Property Get Dom_IEnrollmentColumn_HeaderName() As String
    Dom_IEnrollmentColumn_HeaderName = This.GradeColumn.Number & charUnderScore & This.ClassColumn.Number & charUnderScore & This.GenderColumn.Code
End Property

Private Property Get Dom_IEnrollmentColumn_Value() As String
    Dom_IEnrollmentColumn_Value = This.Value
End Property

