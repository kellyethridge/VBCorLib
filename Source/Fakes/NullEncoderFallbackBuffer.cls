VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullEncoderFallbackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' NullEncoderFallbackBuffer
'
Option Explicit
Implements EncoderFallbackBuffer

Private Function EncoderFallbackBuffer_Equals(Value As Variant) As Boolean

End Function

Private Function EncoderFallbackBuffer_Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional UnknownLowSurrogate As Variant, Optional ByVal Index As Long) As Boolean

End Function

Private Function EncoderFallbackBuffer_GetHashCode() As Long

End Function

Private Function EncoderFallbackBuffer_GetNextChar() As Integer

End Function

Private Function EncoderFallbackBuffer_MovePrevious() As Boolean

End Function

Private Property Get EncoderFallbackBuffer_Remaining() As Long

End Property

Private Sub EncoderFallbackBuffer_Reset()

End Sub

Private Function EncoderFallbackBuffer_ToString() As String

End Function
