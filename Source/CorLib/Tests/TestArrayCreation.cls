VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayCreateInstanceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)

End Sub



Public Sub TestCreateLong()
    Dim a() As Long
    a = NewLongs(1, 2, 3, 4, 5)
    Dim i As Long
    For i = 0 To 4
        Assert.AreEqual i + 1, a(i)
    Next i
End Sub

Public Sub TestEmptyArray()
    Dim a() As Long
    a = cArray.CreateInstance(vbLong)
    Assert.AreEqual -1, UBound(a)
End Sub

Public Sub TestCreateLong2()
    Dim a() As Long
    Dim i As Long
    
    a = NewLongs(89, 191, 409, 907, 1973, 4327, 9461, 20743, 45481, 99607, 218287)
    Assert.AreEqual 10, UBound(a), "NewArray should return 11 elements."
    
    For i = 0 To UBound(a)
        Assert.AreEqual Choose(i + 1, 89, 191, 409, 907, 1973, 4327, 9461, 20743, 45481, 99607, 218287), a(i)
    Next i
End Sub

Public Sub TestCreateString()
    Dim a() As String
    Dim i As Long
    
    a = NewStrings("hi", "bye", "hi again")
    Assert.AreEqual 3, cArray.GetLength(a)
    
    For i = 0 To UBound(a)
        Assert.AreEqual Choose(i + 1, "hi", "bye", "hi again"), a(i)
    Next i
End Sub

Public Sub TestGlobalCreation()
    Dim a() As String
    Dim i As Long
    
    a = NewArray(vbString, 1, 2, 3, 4)
    Assert.That a, Iz.Not.Empty
    Assert.IsFalse cArray.IsNull(a)
    Assert.AreEqual 4, cArray.GetLength(a)
    
    For i = 0 To UBound(a)
        Assert.AreEqual Choose(i + 1, "1", "2", "3", "4"), a(i)
    Next i
End Sub

Public Sub TestCreateEmptyByte()
    Dim b() As Byte
    b = cArray.CreateInstance(vbByte, 0)
    Assert.IsFalse cArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(b)
End Sub

Public Sub TestCreate4Byte()
    Dim b() As Byte
    b = cArray.CreateInstance(vbByte, 4)
    Assert.IsFalse cArray.IsNull(b)
    Assert.AreEqual 4, cArray.GetLength(b)
End Sub

Public Sub TestEmptyNewArray()
    Dim l() As Long
    l = NewArray(vbLong)
    Assert.IsFalse cArray.IsNull(l)
    Assert.AreEqual 0, cArray.GetLength(l)
End Sub

Public Sub TestNewBytes()
    Dim b() As Byte
    b = Cor.NewBytes
    Assert.IsFalse cArray.IsNull(b), "b should not be null."
    Assert.AreEqual 0, cArray.GetLength(b), "b should be zero length."
    
    b = Cor.NewBytes(1, 2, 3)
    Assert.AreEqual 3, cArray.GetLength(b), "b should be 3 bytes."
    Assert.AreEqual 1, b(0)
    Assert.AreEqual 2, b(1)
    Assert.AreEqual 3, b(2)
End Sub

Public Sub TestNewVariants()
    Dim v() As Variant
    v = Cor.NewVariants
    Assert.IsFalse cArray.IsNull(v), "v should not be null."
    Assert.AreEqual 0, cArray.GetLength(v), "v should be zero length."
    
    v = Cor.NewVariants(1, 2, 3)
    Assert.AreEqual 3, cArray.GetLength(v), "v should be 3 elements."
    Assert.AreEqual 1, v(0)
    Assert.AreEqual 2, v(1)
    Assert.AreEqual 3, v(2)
End Sub
