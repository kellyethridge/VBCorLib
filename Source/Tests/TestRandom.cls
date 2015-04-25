VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRandom
'
Option Explicit
Implements ICategorizable

Private r As Random

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set r = New Random
End Sub

Public Sub TestNextLong()
    Dim n As Long
    Dim i As Long
    
    For i = 1 To 100
        n = -1
        n = r.NextLong
        Assert.That -1, Iz.Not.EqualTo(n), "Random number should be generated."
    Next i
End Sub

Public Sub TestNextLongWithMaxValue()
    Dim n As Long
    Dim i As Long
    
    For i = 1 To 1000
        n = r.NextLong(100)
        Assert.IsTrue n < 100, "Random number should be less than 100."
    Next i
End Sub

Public Sub TestInvalidMaxValue()
    On Error GoTo errTrap
    r.NextLong -1
    On Error GoTo 0
    Assert.Fail "An error should be raised."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
End Sub

Public Sub TestNextRange()
    Dim n As Long
    Dim i As Long
    
    For i = 1 To 1000
        n = r.NextRange(10, 100)
        Assert.IsTrue (n >= 10 And n < 100), "Random number is out of range. " & n
    Next i
End Sub

Public Sub TestInvalidNextRangeArguments()
    On Error GoTo errTrap
    r.NextRange 10, 9
    On Error GoTo 0
    Assert.Fail "An error should be raised."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
End Sub

Public Sub TestNextBytes()
    Dim b(100) As Byte
    Dim i As Long
    Dim c As Long
    
    r.NextBytes b
    For i = LBound(b) To UBound(b)
        If b(i) <> 0 Then c = c + 1
    Next i
    Assert.IsTrue c > 0, "Some of the bytes should be changed."
End Sub

Public Sub TestNextDouble()
    Dim d As Double
    Dim i As Long
    
    For i = 1 To 100
        d = -1
        d = r.NextDouble
        Assert.IsTrue d >= 0# And d < 1#, "Random number should be between 0.0 and less than 1.0"
    Next i
End Sub
