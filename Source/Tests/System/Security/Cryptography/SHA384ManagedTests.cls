VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA384ManagedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Test.Use "", "38b060a751ac9638 4cd9327eb1b1e36a 21fdb71114be0743 4c0cc7bf63f6e1da 274edebfe76f65fb d51ad2f14898b95b"
    Test.Use "abc", "cb00753f45a35e8b b5a03d699ac65007 272c32ab0eded163 1a8b605a43ff5bed 8086072ba1e7cc23 58baeca134c825a7"
    Test.Use "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "3391fdddfc8dc739 3707a65b1b470939 7cf8b1d162af05ab fe8f450de5f36bc6 b0455a8520bc4e6f 5fe95b1fe3c8452b"
    Test.Use "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", "09330c33f71147e8 3d192fc782cd1b47 53111b173b3b05d2 2fa08086e3b0f712 fcc7c71a557e2db9 66c3e9fa91746039"
End Sub

Public Sub Constructor_WhenCalled_SetsDefaultValues()
    Dim Provider As New SHA384Managed
    
    Assert.That Provider.HashSize, Iz.EqualTo(384)
    Assert.That Provider.InputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.OutputBlockSize, Iz.EqualTo(1)
    Assert.That Provider.CanReuseTransform, Iz.True
    Assert.That Provider.CanTransformMultipleBlocks, Iz.True
End Sub

Public Sub ComputeHash_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA384Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ComputeHash_WithStream_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Encoding.UTF8.GetBytes(Message), Mode:=AcquireMode)
    Dim Provider As New SHA384Managed
    
    Dim Actual() As Byte
    Actual = Provider.ComputeHash(ms)
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithByteArray_ReturnsExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA384Managed
    Dim Bytes() As Byte
    Bytes = Encoding.UTF8.GetBytes(Message)
    
    Dim ActualHash()    As Byte
    Dim ActualInput()   As Byte
    ActualInput = Provider.TransformFinalBlock(Bytes, 0, CorArray.Length(Bytes))
    ActualHash = Provider.hash
    
    Dim Expected() As Byte
    Expected = HexToBytes(Digest)
    Assert.That ActualInput, Iz.EqualTo(Bytes)
    Assert.That ActualHash, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_With1ByteAtATime_ComputesExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA384Managed
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
    Assert.That Provider.hash, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_With5BytesAtATime_ComputesExpected(ByVal Message As String, ByVal Digest As String)
    Dim Provider As New SHA384Managed
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
    Assert.That Provider.hash, Iz.EqualTo(Expected)
End Sub

Public Sub Hash_BeforeTransformFinalBlockIsCalled_ThrowsInvalidOperationException()
    Dim Provider As New SHA384Managed
    Dim Bytes(0) As Byte
    Provider.TransformBlock Bytes, 0, 1, Bytes, 0
    
    On Error Resume Next
    Dim Actual() As Byte
    Actual = Provider.hash
    
    AssertInvalidOperationException Err
End Sub






