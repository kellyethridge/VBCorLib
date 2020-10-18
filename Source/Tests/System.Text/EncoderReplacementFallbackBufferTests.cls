VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderReplacementFallbackBufferTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncoderReplacementFallbackBufferTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Fallback_WhenCalledForFirstTime_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = MakeBuffer.Fallback(0)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Fallback_WhenCalledForFirstTime_ReturnsTrue2()
    Dim Actual As Boolean
    
    Actual = MakeBuffer.Fallback(&HD800, &HDC00)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Fallback_WithEmptyReplacementString_ReturnsFalse()
    Dim Actual As Boolean
    
    Actual = MakeBuffer("").Fallback(0)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Fallback_WithEmptyReplacementString_ReturnsFalse2()
    Dim Actual As Boolean
    
    Actual = MakeBuffer("").Fallback(&HD800, &HDC00)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub GetNextChar_WhenCalledAfterFallback_ReturnsFallbackChar()
    Dim b As EncoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    b.Fallback 0
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(Asc("?"))
End Sub

Public Sub GetNextChar_WhenCalledAfterFallback_ReturnsFallbackChar2()
    Dim b As EncoderReplacementFallbackBuffer
    Dim Actual1 As Integer
    Dim Actual2 As Integer
    Set b = MakeBuffer
    
    b.Fallback &HD800, &HDC00
    Actual1 = b.GetNextChar
    Actual2 = b.GetNextChar
    
    Assert.That Actual1, Iz.EqualTo(Asc("?"))
    Assert.That Actual2, Iz.EqualTo(Asc("?"))
End Sub

Public Sub GetNextChar_WhenFallbackNotCalled_ReturnsNullChar()
    Dim b As EncoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetNextChar_WhenFallbackFinished_ReturnsNullChar()
    Dim b As EncoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    b.Fallback 0
    Actual = b.GetNextChar
    Actual = b.GetNextChar
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetNextChar_WhenFallbackFinished_ReturnsNullChar2()
    Dim b As EncoderReplacementFallbackBuffer
    Dim Actual As Integer
    Set b = MakeBuffer
    
    b.Fallback &HD800, &HDC00
    Actual = b.GetNextChar
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
    Dim b As EncoderReplacementFallbackBuffer
    Set b = MakeBuffer
    
    Dim Actual As Boolean
    b.Fallback 0
    b.GetNextChar
    Actual = b.MovePrevious
    
    Assert.That Actual, Iz.True
End Sub

Public Sub MovePrevious_WithFallbackStarted_ReturnsTrue2()
    Dim b As EncoderReplacementFallbackBuffer
    Set b = MakeBuffer
    
    Dim Actual As Boolean
    b.Fallback &HD800, &HDC00
    b.GetNextChar
    Actual = b.MovePrevious
    
    Assert.That Actual, Iz.True
End Sub

Public Sub MovePrevious_WhenMovedPrevious_ReturnsFallback()
    Dim b As EncoderFallbackBuffer
    Set b = MakeBuffer
    
    Dim Actual As Integer
    b.Fallback 0
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
    Dim b As EncoderFallbackBuffer
    Dim Actual As Long
    Set b = MakeBuffer
    
    b.Fallback 0
    Actual = b.Remaining
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Remaining_WithFallbackStarted_ReturnsTwo()
    Dim b As EncoderFallbackBuffer
    Dim Actual As Long
    Set b = MakeBuffer
    
    b.Fallback &HD800, &HDC00
    Actual = b.Remaining
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub Fallback_WhenCurrentFallbackNotFinished_ThrowsArgumentException()
    Dim b As EncoderFallbackBuffer
    Set b = MakeBuffer
    
    On Error Resume Next
    b.Fallback 12
    b.Fallback 12
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub Fallback_WhenCurrentFallbackNotFinished_ThrowsArgumentException2()
    Dim b As EncoderFallbackBuffer
    Set b = MakeBuffer
    
    On Error Resume Next
    b.Fallback &HD800, &HDC00
    b.Fallback &HD800, &HDC00
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub Fallback_WithInvalidHighSurrogateAndLowSurrogate_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MakeBuffer.Fallback 0, &HDC00
    
    AssertArgumentOutOfRangeException Err, "UnknownCharOrHighSurrogate"
End Sub

Public Sub Fallback_WithValidHighSurrogateAndInvalidLowSurrogate_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MakeBuffer.Fallback &HD800, 0
    
    AssertArgumentOutOfRangeException Err, "UnknownLowSurrogate"
End Sub

Private Function MakeBuffer(Optional ByRef Replacement As String = "?") As EncoderReplacementFallbackBuffer
    Set MakeBuffer = Cor.NewEncoderReplacementFallback(Replacement).CreateFallbackBuffer
End Function
