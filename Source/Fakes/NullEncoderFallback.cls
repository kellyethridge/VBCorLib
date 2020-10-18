VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullEncoderFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' NullEncoderFallback
'
Option Explicit
Implements EncoderFallback

Private Function EncoderFallback_CreateFallbackBuffer() As CorLib.EncoderFallbackBuffer
    Set EncoderFallback_CreateFallbackBuffer = New NullEncoderFallbackBuffer
End Function

Private Function EncoderFallback_Equals(Value As Variant) As Boolean

End Function

Private Function EncoderFallback_GetHashCode() As Long

End Function

Private Property Get EncoderFallback_MaxCharCount() As Long

End Property

Private Function EncoderFallback_ToString() As String

End Function
