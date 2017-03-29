VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderFallbackExceptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncoderFallbackExceptionTests
'
Option Explicit

Public Sub Constructor_WithMessage_StoresMessage()
    Dim Ex As EncoderFallbackException
    
    Set Ex = Cor.NewEncoderFallbackException("Hello")
    
    Assert.That Ex.Message, Iz.EqualTo("Hello")
    Assert.That Ex.CharUnknown, Iz.EqualTo(0)
    Assert.That Ex.CharUnknownHigh, Iz.EqualTo(0)
    Assert.That Ex.CharUnknownLow, Iz.EqualTo(0)
    Assert.That Ex.Index, Iz.EqualTo(0)
End Sub

Public Sub Constructor_WithCharUnknownAndIndex_StoresParameters()
    Dim Ex As EncoderFallbackException
    
    Set Ex = Cor.NewEncoderFallbackException2("Hello", 1, 2)
    
    Assert.That Ex.Message, Iz.EqualTo("Hello")
    Assert.That Ex.CharUnknown, Iz.EqualTo(1)
    Assert.That Ex.CharUnknownHigh, Iz.EqualTo(0)
    Assert.That Ex.CharUnknownLow, Iz.EqualTo(0)
    Assert.That Ex.Index, Iz.EqualTo(2)
End Sub

Public Sub Constructor_WithCharUnknownHighLowAndIndex_StoresParameters()
    Dim Ex As EncoderFallbackException
    
    Set Ex = Cor.NewEncoderFallbackException3("Hello", 1, 2, 3)
    
    Assert.That Ex.Message, Iz.EqualTo("Hello")
    Assert.That Ex.CharUnknown, Iz.EqualTo(0)
    Assert.That Ex.CharUnknownHigh, Iz.EqualTo(1)
    Assert.That Ex.CharUnknownLow, Iz.EqualTo(2)
    Assert.That Ex.Index, Iz.EqualTo(3)
End Sub

Public Sub Serialization_WithMessage_SerializesInstance()
    Dim Ex As EncoderFallbackException
    Set Ex = Cor.NewEncoderFallbackException("Hello")
    
    Dim Actual As EncoderFallbackException
    Set Actual = Persist(Ex)
    
    Assert.That Actual.Message, Iz.EqualTo("Hello")
    Assert.That Actual.CharUnknown, Iz.EqualTo(0)
    Assert.That Actual.CharUnknownHigh, Iz.EqualTo(0)
    Assert.That Actual.CharUnknownLow, Iz.EqualTo(0)
    Assert.That Actual.Index, Iz.EqualTo(0)
End Sub
