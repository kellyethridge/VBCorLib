VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpanTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpan
'
Option Explicit
Implements ITestCaseSource

Private ts As TimeSpan


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "FromMilliseconds_WithValues_ReturnsTimeSpanOfMilliseconds"
            Test.Use 0
            Test.Use 922337203685477.5807@
            Test.Use -922337203685477.5807@
        
        Case "TotalMilliseconds_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromTicks(5000), 0.5@
            Test.Use TimeSpan.FromMilliseconds(1000), 1000
            Test.Use TimeSpan.FromSeconds(10), 10000
            Test.Use TimeSpan.FromHours(1), 3600000
            Test.Use TimeSpan.FromDays(1), 86400000
        
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use 1, 1, True
            Test.Use 2, 1, False
            Test.Use 1, 2, False
        
        Case "Equals_WithNonTimeSpanValue_ReturnsFalse"
            Test.Use 1
            Test.Use Nothing
            Test.Use Empty
            Test.Use Null
            Test.Use New Collection
        
        Case "Add_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromHours(4), TimeSpan.FromHours(5), TimeSpan.FromHours(9)
            Test.Use TimeSpan.FromMinutes(30), TimeSpan.FromSeconds(15), NewTimeSpan(Minutes:=30, Seconds:=15)
            Test.Use TimeSpan.FromHours(4), TimeSpan.FromHours(-5), TimeSpan.FromHours(-1)
            Test.Use TimeSpan.FromHours(4), Nothing, TimeSpan.FromHours(4)
            
        Case "Subtract_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromHours(5), TimeSpan.FromHours(4), TimeSpan.FromHours(1)
            Test.Use TimeSpan.FromMinutes(30), TimeSpan.FromSeconds(15), NewTimeSpan(Minutes:=29, Seconds:=45)
            Test.Use TimeSpan.FromHours(-1), TimeSpan.FromHours(-4), TimeSpan.FromHours(3)
            Test.Use TimeSpan.FromHours(1), Nothing, TimeSpan.FromHours(1)
            
        Case "Duration_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromHours(-1), TimeSpan.FromHours(1)
            Test.Use TimeSpan.FromHours(1), TimeSpan.FromHours(1)
            
        Case "Negate_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromHours(-1), TimeSpan.FromHours(1)
            Test.Use TimeSpan.FromHours(1), TimeSpan.FromHours(-1)
            
        Case "CompareTo_WithValues_ReturnsExpected"
            Test.Use TimeSpan.FromHours(5), 0
            Test.Use TimeSpan.FromHours(4), 1
            Test.Use TimeSpan.FromHours(6), -1
            Test.Use Nothing, 1
            
        Case "ToString_WhenCalled_ReturnsExpected"
            Test.Use TimeSpan.Zero, "00:00:00"
            Test.Use TimeSpan.FromDays(1), "1.00:00:00"
            Test.Use TimeSpan.FromHours(1), "01:00:00"
            Test.Use TimeSpan.FromMinutes(1), "00:01:00"
            Test.Use TimeSpan.FromSeconds(1), "00:00:01"
            Test.Use TimeSpan.FromMilliseconds(1), "00:00:00.0010000"
            Test.Use TimeSpan.FromTicks(1), "00:00:00.0000001"
            Test.Use TimeSpan.FromTicks(10001), "00:00:00.0010001"
            Test.Use NewTimeSpan(102, 3, 4, 5, 10), "102.03:04:05.0100000"
            Test.Use NewTimeSpan(-2, -3, -4, -5, -10), "-2.03:04:05.0100000"
        
        Case "GetHashCode_WithValues_ReturnsExpected"
            Test.Use 0, 0, True
            Test.Use 1, 1, True
            Test.Use 0, 1, False
            Test.Use 100000, 100000, True
            Test.Use 100001, 100000, False
            
        Case "TryParse_WithValues_ReturnsExpected"
            Test.Use "0", TimeSpan.Zero
            Test.Use "   1:06:07  ", NewTimeSpan(Hours:=1, Minutes:=6, Seconds:=7)
            Test.Use "22:14:14", NewTimeSpan(Hours:=22, Minutes:=14, Seconds:=14)
            Test.Use "1.04:24:34", NewTimeSpan(1, 4, 24, 34)
            Test.Use "5:06:07.1", NewTimeSpan(Hours:=5, Minutes:=6, Seconds:=7, Milliseconds:=100)
            Test.Use "-7.15:35:41.7777777", TimeSpan.FromTicks(-6609417777777#)
            Test.Use "10", TimeSpan.FromDays(10)
            Test.Use "2:3", NewTimeSpan(Hours:=2, Minutes:=3)
            
        Case "TryParse_WithValues_ReturnsExpected2"
            Test.Use "0", True
            Test.Use "1.2", False
            Test.Use "a", False
            Test.Use "0:00", True
            Test.Use "a0", False
            Test.Use "0a", False
            Test.Use "aa:bb", False
            Test.Use "00:aa", False
            Test.Use "0:00.aa", False
            
        Case "Parse_WithOverflowValues_ThrowsOverflowException"
            Test.Use "1:100"
            Test.Use "1:60"
            Test.Use "24:00"
            Test.Use "1:00:60"
            Test.Use "1:00:100"
            Test.Use "10675200.1:00:00"
            
    End Select
End Sub

Public Sub Constructor_WhenCalled_InitsToZero()
    Dim Span As TimeSpan
    
    Set Span = New TimeSpan
    
    Assert.That Span.TotalMilliseconds, Iz.EqualTo(0)
End Sub

Public Sub Zero_WhenCalled_ReturnsTimeSpanOfZero()
    Dim Span As TimeSpan
    
    Set Span = TimeSpan.Zero
    
    Assert.That Span.TotalMilliseconds, Iz.EqualTo(0)
End Sub

Public Sub TotalMilliseconds_WithValues_ReturnsExpected(ByVal Span As TimeSpan, ByVal Expected As Currency)
    Dim Actual As Currency
    
    Actual = Span.TotalMilliseconds
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TotalSeconds_With2500Milliseconds_Returns2Point5()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromMilliseconds(2500)
    
    Dim Actual As Double
    Actual = Span.TotalSeconds
    
    Assert.That Actual, Iz.EqualTo(2.5)
End Sub

Public Sub TotalMinutes_With2700Seconds_Returns45()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromSeconds(2700)
    
    Dim Actual As Double
    Actual = Span.TotalMinutes
    
    Assert.That Actual, Iz.EqualTo(45)
End Sub

Public Sub TotalHours_With90Minutes_Returns1Point5()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromMinutes(90)
    
    Dim Actual As Double
    Actual = Span.TotalHours
    
    Assert.That Actual, Iz.EqualTo(1.5)
End Sub

Public Sub TotalDays_With36Hours_Returns1Point5()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromHours(36)
    
    Dim Actual As Double
    Actual = Span.TotalDays
    
    Assert.That Actual, Iz.EqualTo(1.5)
End Sub

Public Sub Ticks_With1Second_Returns10Million()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromSeconds(1)
    
    Dim Actual As Variant
    Actual = Span.Ticks
    
    Assert.That Actual, Iz.EqualTo(10000000)
End Sub

Public Sub Milliseconds_With1Point5Seconds_Returns500()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromSeconds(1.5)
    
    Dim Actual As Currency
    Actual = Span.Milliseconds
    
    Assert.That Actual, Iz.EqualTo(500)
End Sub

Public Sub Seconds_With1Point5Seconds_Returns1()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromSeconds(1.5)
    
    Dim Actual As Long
    Actual = Span.Seconds
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Hours_With3Point5Hours_Returns3()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromHours(3.5)
    
    Dim Actual As Long
    Actual = Span.Hours
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub Days_With25Hours_Returns1()
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromHours(25)
    
    Dim Actual As Long
    Actual = Span.Days
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByVal LeftHours As Double, ByVal RightHours As Double, ByVal Expected As Boolean)
    Dim Left    As TimeSpan
    Dim Right   As TimeSpan
    Set Left = TimeSpan.FromHours(LeftHours)
    Set Right = TimeSpan.FromHours(RightHours)
    
    Dim Actual As Boolean
    Actual = Left.Equals(Right)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithNonTimeSpanValue_ReturnsFalse(ByVal Value As Variant)
    Dim Span As IObject
    Set Span = New TimeSpan
    
    Dim Actual As Boolean
    Actual = Span.Equals(Value)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Add_WithValues_ReturnsExpected(ByVal Left As TimeSpan, ByVal Right As TimeSpan, ByVal Expected As TimeSpan)
    Dim Actual As TimeSpan
    
    Set Actual = Left.Add(Right)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Subtract_WithValues_ReturnsExpected(ByVal Left As TimeSpan, ByVal Right As TimeSpan, ByVal Expected As TimeSpan)
    Dim Actual As TimeSpan
    
    Set Actual = Left.Subtract(Right)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Duration_WithValues_ReturnsExpected(ByVal Span As TimeSpan, ByVal Expected As TimeSpan)
    Dim Actual As TimeSpan
    
    Set Actual = Span.Duration
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Negate_WithValues_ReturnsExpected(ByVal Span As TimeSpan, ByVal Expected As TimeSpan)
    Dim Actual As TimeSpan
    
    Set Actual = Span.Negate
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub CompareTo_WithValues_ReturnsExpected(ByVal Value As TimeSpan, ByVal Expected As Long)
    Dim Span As TimeSpan
    Set Span = TimeSpan.FromHours(5)
    
    Dim Actual As Long
    Actual = Span.CompareTo(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CompareTo_WithNonTimeSpanValue_ThrowsArgumentException()
    Dim Span As IComparable
    Set Span = New TimeSpan
    
    On Error GoTo errTrap
    Span.CompareTo 1
    
errTrap:
    AssertArgumentException Err, "Value"
End Sub

Public Sub CompareTo_WhenSpanIsZeroAndValueIsNothing_Returns0()
    Dim Span As IComparable
    Set Span = TimeSpan.Zero
    
    Dim Actual As Long
    Actual = Span.CompareTo(Nothing)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub ToString_WhenCalled_ReturnsExpected(ByVal Span As TimeSpan, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Span.ToString
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WithValues_ReturnsExpected(ByVal Milliseconds1 As Currency, ByVal Milliseconds2 As Currency, ByVal Expected As Boolean)
    Dim Span1 As TimeSpan
    Dim Span2 As TimeSpan
    Set Span1 = TimeSpan.FromMilliseconds(Milliseconds1)
    Set Span2 = TimeSpan.FromMilliseconds(Milliseconds2)
    
    Dim Actual As Boolean
    Actual = Span1.GetHashCode = Span2.GetHashCode
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TryParse_WithValues_ReturnsExpected(ByVal TextToParse As String, ByVal Expected As TimeSpan)
    Dim Span As TimeSpan
    
    TimeSpan.TryParse TextToParse, Span
    
    Assert.That Span, Equals(Expected)
End Sub

Public Sub TryParse_WithValues_ReturnsExpected2(ByVal TextToParse As String, ByVal Expected As Boolean)
    Dim Span As TimeSpan
    
    Dim Actual As Boolean
    Actual = TimeSpan.TryParse(TextToParse, Span)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Parse_WithValidSource_ReturnsTimeSpan()
    Const Source As String = "1:00"
    
    Dim Actual As TimeSpan
    Set Actual = TimeSpan.Parse(Source)
    
    Dim Expected As TimeSpan
    Set Expected = TimeSpan.FromHours(1)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Parse_WithInvalidSource_ThrowsFormatError()
    Const InvalidSource As String = "a"
    
    On Error GoTo errTrap
    TimeSpan.Parse InvalidSource
    
errTrap:
    WriteLine Err.Description
    AssertFormatException Err
End Sub

Public Sub Parse_WithOverflowValues_ThrowsOverflowException(ByVal s As String)
    On Error GoTo errTrap
    
    TimeSpan.Parse s
    
errTrap:
    AssertOverflowException Err
End Sub









