VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACTripleDESTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MACTripleDESTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private mKey() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ComputeHash_WhenSettingKeyProperty_ReturnsExpected"
            AddTests Test
            
        Case "ComputeHash_WhenConstructingWithKey_ReturnsExpected"
            AddTests Test
            
        Case "ComputeHash_WithSpecifiedSubArray_ReturnsExpected"
            AddTests Test
            
        Case "ComputeHash_WithStream_ReturnsExpected"
            AddTests Test
            
        Case "TransformFinalBlock_WhenCalled_ComputesHash"
            AddTests Test
            
        Case "TransformBlock_WhenCalled_ComputesHash"
            AddTests Test
            
    End Select
End Sub

Private Sub AddTests(ByVal Test As TestCaseCollector)
    Test.Use ByteArray(7), NewBytes(&H86, &HE9, &H65, &HBD, &H1E, &HC4, &H44, &H61)
    Test.Use ByteArray(8), NewBytes(&H86, &HE9, &H65, &HBD, &H1E, &HC4, &H44, &H61)
    Test.Use ByteArray(14), NewBytes(&H23, &HD6, &H92, &HA0, &H80, &H6E, &HC9, &H30)
    Test.Use ByteArray(48), NewBytes(&HA3, &HE, &H34, &H26, &H8B, &H49, &HEF, &H49)
End Sub

Public Sub FixtureSetup()
    mKey = HexToBytes("0123456789abcdef23456789abcdef010123456789abcdef")
End Sub

Public Sub Constructor_WithNullKey_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewMACTripleDES NullBytes
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub Constructor_WithKeyAndName_CreatesInstance()
    Cor.NewMACTripleDES mKey, "TripleDES"
End Sub

Public Sub Constructor_WithInvalidAlgorithmName_ThrowsInvalidCastException()
    On Error Resume Next
    
    Cor.NewMACTripleDES mKey, "DES"
    
    AssertInvalidCastException Err
End Sub

Public Sub Invariants()
    Dim csp As New MACTripleDES
    Assert.That csp.CanReuseTransform, Iz.True, "CanReuseTransform"
    Assert.That csp.CanTransformMultipleBlocks, Iz.True, "CanTransformMultipleBlocks"
    Assert.That csp.HashSize, Iz.EqualTo(64), "HashSize"
    Assert.That csp.InputBlockSize, Iz.EqualTo(1), "InputBlockSize"
    Assert.That csp.OutputBlockSize, Iz.EqualTo(1), "OutputBlockSize"
    Assert.That csp.ToString, Iz.EqualTo("CorLib.MACTripleDES"), "ToString"
    Assert.That csp.Key, Iz.Not.Empty
End Sub

Public Sub ComputeHash_WhenSettingKeyProperty_ReturnsExpected(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim csp As New MACTripleDES
    csp.Key = mKey
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Msg)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WhenConstructingWithKey_ReturnsExpected(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim csp As MACTripleDES
    Set csp = Cor.NewMACTripleDES(mKey)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Msg)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WithSpecifiedSubArray_ReturnsExpected(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim csp As New MACTripleDES
    csp.Key = mKey
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Msg, LBound(Msg), CorArray.Length(Msg))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WithStream_ReturnsExpected(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim csp As New MACTripleDES
    Dim ms As MemoryStream
    csp.Key = mKey
    Set ms = Cor.NewMemoryStream(Msg)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(ms)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WhenCalled_ComputesHash(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim csp As New MACTripleDES
    csp.Key = mKey
    
    Dim Actual() As Byte
    Actual = csp.TransformFinalBlock(Msg, LBound(Msg), CorArray.Length(Msg))
    
    Assert.That Actual, Iz.EqualTo(Msg)
    Assert.That csp.Hash, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_WhenCalled_ComputesHash(ByRef Msg() As Byte, ByRef Expected() As Byte)
    Dim Copy(0 To 1) As Byte
    Dim csp As New MACTripleDES
    csp.Key = mKey
    
    Dim i As Long
    For i = LBound(Msg) To UBound(Msg) - 1
        csp.TransformBlock Msg, i, 1, Copy, 0
    Next
    csp.TransformFinalBlock Msg, UBound(Msg), 1
    
    Assert.That csp.Hash, Iz.EqualTo(Expected)
End Sub

Private Function ByteArray(ByVal Size As Long) As Byte()
    Dim Arr() As Byte
    ReDim Arr(0 To Size - 1)
    ByteArray = Arr
End Function
