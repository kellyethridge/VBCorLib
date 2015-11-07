VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA1ManagedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   SHA1ManagedTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ComputeHash_WithByteArray_ReturnsExpected"
            AddHashTests Test
            
        Case "ComputeHash_WithStream_ReturnsExpected"
            AddHashTests Test
            
        Case "TransformFinalBlock_WithByteArray_ReturnsExpected"
            AddHashTests Test
            
        Case "TransformBlock_With1ByteAtATime_ComputesExpected"
            AddHashTests Test
            
        Case "TransformBlock_With5BytesAtATime_ComputesExpected"
            AddHashTests Test
            
    End Select
End Sub

Private Sub AddHashTests(ByVal Test As TestCaseCollector)
    Test.Use "", "da39a3ee 5e6b4b0d 3255bfef 95601890 afd80709"
    Test.Use "abc", "a9993e36 4706816a ba3e2571 7850c26c 9cd0d89d"
    Test.Use "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "84983e44 1c3bd26e baae4aa1 f95129e5 e54670f1"
    Test.Use "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", "a49b2446 a02c645b f419f995 b6709125 3a04a259"
End Sub

Public Sub Constructor_WhenCalled_SetsDefaultValues()
    Dim Provider As New SHA1Managed
    
    Assert.That Provider.HashSize, Iz.EqualTo(160)
    Assert.That Provider.InputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.OutputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.CanReuseTransform, Iz.True
    Assert.That Provider.CanTransformMultipleBlocks, Iz.True
End Sub

Public Sub ComputeHash_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA1Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WithStream_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Encoding.UTF8.GetBytes(Message), Mode:=AcquireMode)
    Dim Provider As New SHA1Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(ms)
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA1Managed
    Dim Bytes() As Byte
    Bytes = Encoding.UTF8.GetBytes(Message)
    
    Dim ActualHash()    As Byte
    Dim ActualInput()   As Byte
    ActualInput = Provider.TransformFinalBlock(Bytes, 0, CorArray.Length(Bytes))
    ActualHash = Provider.Hash
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That ActualInput, Iz.EqualTo(Bytes)
    Assert.That ActualHash, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_With1ByteAtATime_ComputesExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA1Managed
    Dim PT() As Byte
    PT = Encoding.UTF8.GetBytes(Message)
    Dim Output(0) As Byte
    
    Dim i As Long
    For i = 0 To UBound(PT)
        Assert.That Provider.TransformBlock(PT, i, 1, Output, 0), Iz.EqualTo(1)
        Assert.That Output(0), Iz.EqualTo(PT(i))
    Next i
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Provider.TransformFinalBlock(PT, UBound(PT) + 1, 0), Iz.Empty
    Assert.That Provider.Hash, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_With5BytesAtATime_ComputesExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA1Managed
    Dim PT() As Byte
    PT = Encoding.UTF8.GetBytes(Message)
    Dim Output(4) As Byte
    
    Dim i As Long
    Do While i <= (UBound(PT) - 5)
        Assert.That Provider.TransformBlock(PT, i, 5, Output, 0), Iz.EqualTo(5)
        Dim j As Long
        For j = 0 To 4
            Assert.That Output(j), Iz.EqualTo(PT(i + j))
        Next j

        i = i + 5
    Loop
    Provider.TransformFinalBlock PT, i, UBound(PT) - i + 1
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Provider.Hash, Iz.EqualTo(Expected)
End Sub

Public Sub Hash_BeforeTransformFinalBlockIsCalled_ThrowsInvalidOperationException()
    Dim Provider As New SHA1Managed
    Dim Bytes(0) As Byte
    Provider.TransformBlock Bytes, 0, 1, Bytes, 0
    
    On Error Resume Next
    Dim Actual() As Byte
    Actual = Provider.Hash
    
    AssertInvalidOperationException Err
End Sub




