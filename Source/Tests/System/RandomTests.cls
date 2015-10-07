VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RandomTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub NextLong_WhenCalled_ReturnsRandomPositiveNumber()
    Dim Rand As New Random
    Dim i As Long
    
    For i = 1 To 100
        Dim Actual As Long
        Actual = Rand.NextLong
        Assert.That Actual, Iz.GreaterThanOrEqualTo(0)
    Next
End Sub

Public Sub NextLong_WithMaxValue_ReturnsValuesLessThanMaxValue()
    Dim Rand As New Random
    Dim i As Long
    
    For i = 1 To 100
        Dim Actual As Long
        Actual = Rand.NextLong(100)
        Assert.That Actual, Iz.GreaterThanOrEqualTo(0).AndAlso.LessThan(100)
    Next i
End Sub

Public Sub MaxValue_WithValueLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Rand As New Random
    
    On Error Resume Next
    Rand.NextLong -1
    
    AssertArgumentOutOfRangeException Err, "MaxValue"
End Sub

Public Sub NextRange_WhenCalled_ReturnsValuesWithinRange()
    Dim Rand As New Random
    Dim i As Long
    
    For i = 1 To 100
        Dim Actual As Long
        Actual = Rand.NextRange(10, 100)
        Assert.That Actual, Iz.GreaterThanOrEqualTo(10).AndAlso.LessThan(100)
    Next i
End Sub

Public Sub NextRange_WithNegativeValues_ReturnsValuesWithinRange()
    Dim Rand As New Random
    Dim i As Long
    
    For i = 1 To 100
        Dim Actual As Long
        Actual = Rand.NextRange(-100, -10)
        Assert.That Actual, Iz.GreaterThanOrEqualTo(-100).AndAlso.LessThan(-10)
    Next i
End Sub

Public Sub NextRange_WithMinValueEqualMaxValue_ReturnsMinValue()
    Dim Rand As New Random
    
    Dim Actual As Long
    Actual = Rand.NextRange(5000, 5000)
    
    Assert.That Actual, Iz.EqualTo(5000)
End Sub

Public Sub NextRange_WithMinValueGreaterThanMaxValue_ThrowsArgumentOutOfRangeException()
    Dim Rand As New Random
    
    On Error Resume Next
    Rand.NextRange 100, 99
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub NextBytes_WhenCalled_FillsArrayWithValues()
    Dim Rand As New Random
    
    Dim Actual(10 To 100) As Byte
    Rand.NextBytes Actual
    
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub NextBytes_WhenCalled_FillsWithRandomValues()
    Dim Rand As New Random
    
    Dim Bytes1(10) As Byte
    Dim Bytes2(10) As Byte
    Rand.NextBytes Bytes1
    Rand.NextBytes Bytes2
    
    Assert.That Bytes1, Iz.Not.EqualTo(Bytes2)
End Sub

Public Sub NextBytes_WithSameSeed_ReturnsSamePattern()
    Dim Rand As Random
    Dim Bytes1(10) As Byte
    Dim Bytes2(10) As Byte
    
    Set Rand = Cor.NewRandom(-5)
    Rand.NextBytes Bytes1
    
    Set Rand = Cor.NewRandom(-5)
    Rand.NextBytes Bytes2
    
    Assert.That Bytes1, Iz.EqualTo(Bytes2)
End Sub

Public Sub NextBytes_WithNullArray_ThrowsArgumentNullException()
    Dim Rand As New Random
    Dim NullBytes() As Byte
    
    On Error Resume Next
    Rand.NextBytes NullBytes
    
    AssertArgumentNullException Err, "Bytes"
End Sub

