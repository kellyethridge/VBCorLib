VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA512ManagedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   SHA512ManagedTests
'
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
    Test.Use "", "cf83e1357eefb8bd f1542850d66d8007 d620e4050b5715dc 83f4a921d36ce9ce 47d0d13c5d85f2b0 ff8318d2877eec2f 63b931bd47417a81 a538327af927da3e"
    Test.Use "abc", "ddaf35a193617aba cc417349ae204131 12e6fa4e89a97ea2 0a9eeee64b55d39a 2192992a274fc1a8 36ba3c23a3feebbd 454d4423643ce80e 2a9ac94fa54ca49f"
    Test.Use "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "204a8fc6dda82f0a 0ced7beb8e08a416 57c16ef468b228a8 279be331a703c335 96fd15c13b1b07f9 aa1d3bea57789ca0 31ad85c7a71dd703 54ec631238ca3445"
    Test.Use "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", "8e959b75dae313da 8cf4f72814fc143f 8f7779c6eb9f7fa1 7299aeadb6889018 501d289e4900f7e4 331b99dec4b5433a c7d329eeb6dd2654 5e96e55b874be909"
End Sub

Public Sub Constructor_WhenCalled_SetsDefaultValues()
    Dim Provider As New SHA512Managed
    
    Assert.That Provider.HashSize, Iz.EqualTo(512)
    Assert.That Provider.InputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.OutputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.CanReuseTransform, Iz.True
    Assert.That Provider.CanTransformMultipleBlocks, Iz.True
End Sub

Public Sub ComputeHash_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA512Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WithStream_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Encoding.UTF8.GetBytes(Message), Mode:=AcquireMode)
    Dim Provider As New SHA512Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(ms)
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA512Managed
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
    Dim Provider As New SHA512Managed
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
    Dim Provider As New SHA512Managed
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
    Dim Provider As New SHA512Managed
    Dim Bytes(0) As Byte
    Provider.TransformBlock Bytes, 0, 1, Bytes, 0
    
    On Error Resume Next
    Dim Actual() As Byte
    Actual = Provider.Hash
    
    AssertInvalidOperationException Err
End Sub






