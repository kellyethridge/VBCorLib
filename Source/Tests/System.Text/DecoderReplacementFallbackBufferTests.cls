VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallbackBufferTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DecoderReplacementFallbackBuffer Tests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Fallback_WhenCalledForFirstTime_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = MakeBuffer.Fallback(NewBytes(), 0)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Fallback_WhenEmptyFallbackString_ReturnsFalse()
    Dim Fallback As DecoderReplacementFallback
    Dim Buffer As DecoderReplacementFallbackBuffer
    Set Fallback = Cor.NewDecoderReplacementFallback("")
    Set Buffer = Cor.NewDecoderReplacementFallbackBuffer(Fallback)
    
    Dim Actual As Boolean
    Actual = Buffer.Fallback(NewBytes(), 0)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub GetNextChar_WhenCalledAfterFallback_ReturnsFallbackChar()
    Dim b As DecoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    b.Fallback NewBytes(), 0
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(Asc("?"))
End Sub

Public Sub GetNextChar_WhenFallbackNotCalled_ReturnsNullChar()
    Dim b As DecoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetNextChar_WhenFallbackFinished_ReturnsNullChar()
    Dim b As DecoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    b.Fallback NewBytes(), 0
    Actual = b.GetNextChar
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub MovePrevious_WithFallbackNotStarted_ReturnsFalse()
    Dim Actual As Boolean
    
    Actual = MakeBuffer.MovePrevious
    
    Assert.That Actual, Iz.False
End Sub

Public Sub MovePrevious_WithFallbackStarted_ReturnsTrue()
    Dim b As DecoderFallbackBuffer
    Set b = MakeBuffer
    
    Dim Actual As Boolean
    b.Fallback NewBytes(), 0
    b.GetNextChar
    Actual = b.MovePrevious
    
    Assert.That Actual, Iz.True
End Sub

Public Sub MovePrevious_WhenMovedPrevious_ReturnsFallback()
    Dim b As DecoderFallbackBuffer
    Set b = MakeBuffer
    
    Dim Actual As Integer
    b.Fallback NewBytes(), 0
    b.GetNextChar
    b.MovePrevious
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(Asc("?"))
End Sub

Public Sub Remaining_WithFallbackNotStarted_ReturnsZero()
    Dim Actual As Long
    
    Actual = MakeBuffer.Remaining
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Remaining_WithFallbackStarted_ReturnsOne()
    Dim b As DecoderFallbackBuffer
    Dim Actual As Long
    Set b = MakeBuffer
    
    b.Fallback NewBytes(), 0
    Actual = b.Remaining
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Fallback_WhenCurrentFallbackNotFinished_ThrowsArgumentException()
    Dim b As DecoderFallbackBuffer
    Set b = MakeBuffer
    
    On Error Resume Next
    b.Fallback NewBytes(), 0
    b.Fallback NewBytes(), 0
    
    AssertArgumentException Err, "BytesUnknown"
End Sub



Private Function MakeBuffer() As DecoderReplacementFallbackBuffer
    Dim Decoder As New DecoderReplacementFallback
    Set MakeBuffer = Cor.NewDecoderReplacementFallbackBuffer(Decoder)
End Function
