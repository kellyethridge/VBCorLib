VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestToBase64Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestToBase64Transform
'
Option Explicit
Implements ICategorizable

Private Trans As ToBase64Transform

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set Trans = New ToBase64Transform
End Sub

Public Sub TestDefaultPropertyValues()
    Assert.IsTrue Trans.CanReuseTransform, "Can reuse transform."
    Assert.IsFalse Trans.CanTransformMultipleBlocks, "Cannot transform multiple blocks."
    Assert.AreEqual 3, Trans.InputBlockSize, "Input block size should b 3 bytes."
    Assert.AreEqual 4, Trans.OutputBlockSize, "Output block size should be 4 bytes."
End Sub

Public Sub TestTransformBlockDisposedException()
    Trans.Clear
    On Error GoTo errTrap
    Dim b() As Byte
    Trans.TransformBlock b, 0, 0, b, 0
    On Error GoTo 0
    Assert.Fail "An exception should be raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ObjectDisposedException, "Wrong exception type caught."
End Sub

Public Sub TestTransform3Zeros()
    Dim b(2) As Byte
    Dim Output(3) As Byte
    
    Assert.AreEqual 4, Trans.TransformBlock(b, 0, 3, Output, 0)
    AssertOutput Output, 65, 65, 65, 65
End Sub

Public Sub TestTransform123()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3)
    
    Dim Output(3) As Byte
    Assert.AreEqual 4, Trans.TransformBlock(b, 0, 3, Output, 0)
    AssertOutput Output, 65, 81, 73, 68
End Sub

Public Sub TestTransformOffset()
    Dim b() As Byte
    b = NewArray(vbByte, 0, 0, 0, 1, 2, 3, 0, 0, 0)
    
    Dim Output(10) As Byte
    Assert.AreEqual 4, Trans.TransformBlock(b, 3, 3, Output, 4)
    AssertOutput Output, 65, 81, 73, 68, 4
End Sub

Public Sub TestTransformFinal3Zeros()
    Dim b(2) As Byte
    Dim Output() As Byte
    
    Output = Trans.TransformFinalBlock(b, 0, 3)
    
    AssertOutput Output, 65, 65, 65, 65
End Sub

Public Sub TestTransformFinal2Zeros()
    Dim b(2) As Byte
    Dim Output() As Byte
    
    Output = Trans.TransformFinalBlock(b, 0, 2)
    
    AssertOutput Output, 65, 65, 65, 61
End Sub

Public Sub TestTransformFinal1Zero()
    Dim b(2) As Byte
    Dim Output() As Byte
    
    Output = Trans.TransformFinalBlock(b, 0, 1)
    
    AssertOutput Output, 65, 65, 61, 61
End Sub

Public Sub TestTransformFinal123()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3)
    
    Dim Output() As Byte
    Output = Trans.TransformFinalBlock(b, 0, 3)
    
    AssertOutput Output, 65, 81, 73, 68
End Sub

Public Sub TestTransformFinal12()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2)
    
    Dim Output() As Byte
    Output = Trans.TransformFinalBlock(b, 0, 2)
    
    AssertOutput Output, 65, 81, 73, 61
End Sub

Public Sub TestTransformFinal1()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    
    Dim Output() As Byte
    Output = Trans.TransformFinalBlock(b, 0, 1)
    
    AssertOutput Output, 65, 81, 61, 61
End Sub

Public Sub TestTransformFinalEmptyArray()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    
    Dim Output() As Byte
    Output = Trans.TransformFinalBlock(b, 0, 0)
    
    Assert.IsFalse CorArray.IsNull(Output), "Output array should not be null."
    Assert.AreEqual 0, CorArray.GetLength(Output), "Output array should be empty."
End Sub

Public Sub TestTransformFinalOffset()
    Dim b() As Byte
    b = NewArray(vbByte, 0, 0, 0, 1, 2, 3, 0, 0, 0)
    
    Dim Output() As Byte
    Output = Trans.TransformFinalBlock(b, 3, 3)
    
    AssertOutput Output, 65, 81, 73, 68
End Sub

Private Sub AssertOutput(ByRef Output() As Byte, ByVal v1 As Byte, ByVal v2 As Byte, ByVal v3 As Byte, ByVal v4 As Byte, Optional ByVal Offset As Long = 0)
    Assert.AreEqual v1, Output(Offset), "First byte is wrong."
    Assert.AreEqual v2, Output(Offset + 1), "Second byte is wrong."
    Assert.AreEqual v3, Output(Offset + 2), "Third byte is wrong."
    Assert.AreEqual v4, Output(Offset + 3), "Fourth byte is wrong."
End Sub
