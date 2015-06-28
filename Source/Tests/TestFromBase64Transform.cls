VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFromBase64Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFromBase64Transform
'
Option Explicit
Implements ICategorizable

Private Trans As FromBase64Transform

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set Trans = New FromBase64Transform
End Sub

Public Sub TestDefaultPropertyValues()
    Assert.IsTrue Trans.CanReuseTransform
    Assert.IsFalse Trans.CanTransformMultipleBlocks
    Assert.AreEqual 1, Trans.InputBlockSize
    Assert.AreEqual 3, Trans.OutputBlockSize
End Sub

Public Sub TestTransfromBackTo123()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 81, 73, 68)
    
    Dim original(2) As Byte
    Assert.AreEqual 3, Trans.TransformBlock(b, 0, 4, original, 0)
    Assert.AreEqual 1, original(0), "byte 0 is wrong."
    Assert.AreEqual 2, original(1), "byte 1 is wrong."
    Assert.AreEqual 3, original(2), "byte 2 is wrong."
End Sub

Public Sub TestTransformBackTo123456()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 81, 73, 68, 66, 65, 85, 71)
    
    Dim original(5) As Byte
    Assert.AreEqual 6, Trans.TransformBlock(b, 0, 8, original, 0)
    
    Dim i As Long
    For i = 0 To 5
        Assert.AreEqual i + 1, original(i), "byte " & i & " is wrong."
    Next i
End Sub

Public Sub TestTransformBackTo123456InTwoCalls()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 81, 73, 68, 66, 65, 85, 71)
    
    Dim original(5) As Byte
    Assert.AreEqual 3, Trans.TransformBlock(b, 0, 6, original, 0)
    Assert.AreEqual 1, original(0)
    Assert.AreEqual 2, original(1)
    Assert.AreEqual 3, original(2)
    Assert.AreEqual 0, original(3)
    Assert.AreEqual 0, original(4)
    Assert.AreEqual 0, original(5)
    
    Assert.AreEqual 3, Trans.TransformBlock(b, 6, 2, original, 3)
    Assert.AreEqual 1, original(0)
    Assert.AreEqual 2, original(1)
    Assert.AreEqual 3, original(2)
    Assert.AreEqual 4, original(3)
    Assert.AreEqual 5, original(4)
    Assert.AreEqual 6, original(5)
End Sub

Public Sub TestTransformBackTo123456IgnoringWhiteSpaces()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 13, 81, 73, 9, 68, 66, 32, 10, 65, 85, 71)
    
    Dim original(5) As Byte
    Assert.AreEqual 6, Trans.TransformBlock(b, 0, CorArray.Length(b), original, 0)
    Dim i As Long
    For i = 0 To 5
        Assert.AreEqual i + 1, original(i), "Byte " & i & " is wrong."
    Next i
End Sub

Public Sub TestTranformDoNotIgnoreWhiteSpaces()
    Set Trans = Cor.NewFromBase64Transform(DoNotIgnoreWhiteSpaces)
    
    Dim b() As Byte
    b = NewArray(vbByte, 65, 13, 81, 73, 9, 68, 66, 32, 10, 65, 85, 71)

    Dim original(5) As Byte
    On Error GoTo errTrap
    Trans.TransformBlock b, 0, CorArray.Length(b), original, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception type caught."
End Sub

Public Sub TestTransformWithOneTerminationCharacter()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 81, 73, 61)
    Dim original(3) As Byte
    Assert.AreEqual 2, Trans.TransformBlock(b, 0, 4, original, 0)
    Assert.AreEqual 1, original(0)
    Assert.AreEqual 2, original(1)
    Assert.AreEqual 0, original(2)
End Sub

Public Sub TestTransformWithTwoTerminationCharacters()
    Dim b() As Byte
    b = NewArray(vbByte, 67, 103, 61, 61)
    Dim original(1) As Byte
    Assert.AreEqual 1, Trans.TransformBlock(b, 0, 4, original, 0)
    Assert.AreEqual 10, original(0)
    Assert.AreEqual 0, original(1)
End Sub

Public Sub TestTransformWithThreeTerminationCharacters()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 61, 61, 61)
    Dim original(1) As Byte
    On Error GoTo errTrap
    Trans.TransformBlock b, 0, 4, original, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception type caught."
    
End Sub

Public Sub TestIncrementalTransform()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    b = Encoding.ASCII.GetBytes(Convert.ToBase64String(b))
    
    Dim b2(7) As Byte
    Dim t2 As New FromBase64Transform
    Dim P As Long
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(b)
        P = P + t2.TransformBlock(b, i, 1, b2, P)
        Select Case i
            Case 0 To 2
                For j = 0 To 7
                    Assert.AreEqual 0, b2(j), "Iteration " & i + 1 & ", byte " & j
                Next j
                
            Case 3 To 6
                For j = 0 To 7
                    If j < 3 Then
                        Assert.AreEqual j + 1, b2(j), "Iteration " & i + 1 & ", byte " & j
                    Else
                        Assert.AreEqual 0, b2(j), "Iteration " & i + 1 & ", byte " & j
                    End If
                Next j
                
            Case 7 To 10
                For j = 0 To 7
                    If j < 6 Then
                        Assert.AreEqual j + 1, b2(j), "Iteration " & i + 1 & ", byte " & j
                    Else
                        Assert.AreEqual 0, b2(j), "Iteration " & i + 1 & ", byte " & j
                    End If
                Next j
                
            Case Else
                For j = 0 To 7
                    Assert.AreEqual j + 1, b2(j), "Iteration " & i + 1 & ", byte " & j
                Next j
        End Select
    Next
End Sub

Public Sub TestTransformFinalBackTo123()
    Dim b() As Byte
    b = NewArray(vbByte, 65, 81, 73, 68)
    
    Dim original() As Byte
    original = Trans.TransformFinalBlock(b, 0, 4)
    Assert.AreEqual 3, CorArray.Length(original), "Array length should be 3."
    Assert.AreEqual 1, original(0), "byte 0 is wrong."
    Assert.AreEqual 2, original(1), "byte 1 is wrong."
    Assert.AreEqual 3, original(2), "byte 2 is wrong."
End Sub

Public Sub TestTransformFinalIncrementalToOne()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    b = Encoding.ASCII.GetBytes(Convert.ToBase64String(b))
    
    Dim i As Long
    For i = 0 To 3
        Dim b2() As Byte
        b2 = Trans.TransformFinalBlock(b, 0, i)
        Assert.IsFalse CorArray.IsNull(b2), "b2 should not be null."
        Assert.AreEqual 0, CorArray.Length(b2)
    Next i
    
    b2 = Trans.TransformFinalBlock(b, 0, 4)
    Assert.IsFalse CorArray.IsNull(b2), "b2 should not be null."
    Assert.AreEqual 1, CorArray.Length(b2), "b2 should have 1 element."
    Assert.AreEqual 1, b2(0), "byte 0 should equal 1."
End Sub

Public Sub TestTransformFinal1BytePartialIncrementalToOne()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    b = Encoding.ASCII.GetBytes(Convert.ToBase64String(b))
    
    Dim b2() As Byte
    Dim i As Long
    For i = 1 To 2
        ReDim b2(0)
        Trans.TransformBlock b, 0, 1, b2, 0 ' get partial transform completed
        
        b2 = Trans.TransformFinalBlock(b, 1, i)
        Assert.IsFalse CorArray.IsNull(b2)
        Assert.AreEqual 0, CorArray.Length(b2)
    Next i
    
    ReDim b2(0)
    Trans.TransformBlock b, 0, 1, b2, 0 ' get partial transform completed
    b2 = Trans.TransformFinalBlock(b, 1, 3)
    Assert.AreEqual 1, CorArray.Length(b2)
    Assert.AreEqual 1, b2(0)
End Sub

Public Sub TestTransformFinal2BytePartialIncrementalToOne()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    b = Encoding.ASCII.GetBytes(Convert.ToBase64String(b))
    
    Dim b2() As Byte
    Dim i As Long
    ReDim b2(0)
    Trans.TransformBlock b, 0, 2, b2, 0 ' get partial transform completed
    
    b2 = Trans.TransformFinalBlock(b, 2, 1)
    Assert.IsFalse CorArray.IsNull(b2)
    Assert.AreEqual 0, CorArray.Length(b2)
    
    ReDim b2(0)
    Trans.TransformBlock b, 0, 2, b2, 0 ' get partial transform completed
    b2 = Trans.TransformFinalBlock(b, 2, 2)
    Assert.AreEqual 1, CorArray.Length(b2)
    Assert.AreEqual 1, b2(0)
End Sub

Public Sub TestTransformFinal3BytePartialIncrementalToOne()
    Dim b() As Byte
    b = NewArray(vbByte, 1)
    b = Encoding.ASCII.GetBytes(Convert.ToBase64String(b))
    
    Dim b2() As Byte
    Dim i As Long
    ReDim b2(0)
    Trans.TransformBlock b, 0, 3, b2, 0  ' get partial transform completed
    
    b2 = Trans.TransformFinalBlock(b, 3, 1)
    Assert.AreEqual 1, CorArray.Length(b2)
    Assert.AreEqual 1, b2(0)
End Sub




