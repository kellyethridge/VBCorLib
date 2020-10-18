VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderExceptionFallbackBufferTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncoderExceptionFallbackBufferTests
'
Option Explicit

Public Sub Defaults()
    Dim Buffer As New EncoderExceptionFallbackBuffer
    
    Assert.That Buffer.GetNextChar, Iz.EqualTo(0)
    Assert.That Buffer.MovePrevious, Iz.False
    Assert.That Buffer.Remaining, Iz.EqualTo(0)
End Sub

Public Sub Fallback_WithCharUnknown_ThrowsEncoderFallbackException()
    Dim Buffer As New EncoderExceptionFallbackBuffer
    
    On Error Resume Next
    Buffer.Fallback 1, Index:=2
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.CharUnknown, Iz.EqualTo(1)
    Assert.That Ex.Index, Iz.EqualTo(2)
End Sub

Public Sub Fallback_WithHighLowSurrogates_ThrowsEncoderFallbackException()
    Dim Buffer As New EncoderExceptionFallbackBuffer
    
    On Error Resume Next
    Buffer.Fallback &HD800, &HDC00, 5
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.CharUnknownHigh, Iz.EqualTo(&HD800)
    Assert.That Ex.CharUnknownLow, Iz.EqualTo(&HDC00)
    Assert.That Ex.Index, Iz.EqualTo(5)
End Sub

Public Sub Fallback_WithInvalidHighSurrogate_ThrowsArgumentOutOfRangeException()
    Dim Buffer As New EncoderExceptionFallbackBuffer
    
    On Error Resume Next
    Buffer.Fallback 1, &HDC00
    
    AssertArgumentOutOfRangeException Err, "UnknownCharOrHighSurrogate"
End Sub

Public Sub Fallback_WithInvalidLowSurrogate_ThrowsArgumentOutOfRangeException()
    Dim Buffer As New EncoderExceptionFallbackBuffer
    
    On Error Resume Next
    Buffer.Fallback &HD800, 1
    
    AssertArgumentOutOfRangeException Err, "UnknownLowSurrogate"
End Sub
