VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArrayTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Test BitArray
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Constructor_WithDefaultValue_CreatesInitializedList"
            Test.Use True
            Test.Use False
            
        Case "CopyTo_WithValues_PerformsCopy"
            Test.Use NewLongs(0, 0), NewLongs(1, 2)
            Test.Use NewBytes(0, 0, 0, 0, 0), NewBytes(&H1, 0, 0, 0, &H2)
            Test.Use CorArray.CreateInstance(vbBoolean, 34), NewBooleans(True, False, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, True)
            
        Case "CopyTo_WithValuesAndOffset_PerformatCopy"
            Test.Use NewLongs(0, 0, 0), NewLongs(0, 1, 2)
            Test.Use NewBytes(0, 0, 0, 0, 0, 0), NewBytes(0, &H1, 0, 0, 0, &H2)
            Test.Use CorArray.CreateInstance(vbBoolean, 35), NewBooleans(False, True, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, False, False, False, False, False, False, False, _
                                                                         False, False, False, False, True)
        
    End Select
End Sub

Public Sub Constructor_WhenCalled_CreatesListInitializedToFalse()
    Dim List As BitArray
    
    Set List = Cor.NewBitArray(10)
    
    Assert.That List(0), Iz.False
    Assert.That List(1), Iz.False
    Assert.That List(2), Iz.False
    Assert.That List(3), Iz.False
    Assert.That List(4), Iz.False
    Assert.That List(5), Iz.False
    Assert.That List(6), Iz.False
    Assert.That List(7), Iz.False
    Assert.That List(8), Iz.False
    Assert.That List(9), Iz.False
End Sub

Public Sub Constructor_WithDefaultValue_CreatesInitializedList(ByVal DefaultValue As Boolean)
    Dim List As BitArray
    
    Set List = Cor.NewBitArray(10, DefaultValue)
    
    Assert.That List(0), Iz.EqualTo(DefaultValue)
    Assert.That List(1), Iz.EqualTo(DefaultValue)
    Assert.That List(2), Iz.EqualTo(DefaultValue)
    Assert.That List(3), Iz.EqualTo(DefaultValue)
    Assert.That List(4), Iz.EqualTo(DefaultValue)
    Assert.That List(5), Iz.EqualTo(DefaultValue)
    Assert.That List(6), Iz.EqualTo(DefaultValue)
    Assert.That List(7), Iz.EqualTo(DefaultValue)
    Assert.That List(8), Iz.EqualTo(DefaultValue)
    Assert.That List(9), Iz.EqualTo(DefaultValue)
End Sub

Public Sub Item_WhenSet_ChangedBit()
    Dim List As BitArray
    Set List = Cor.NewBitArray(10)
    
    List(2) = True
    
    Assert.That List(1), Iz.False
    Assert.That List(2), Iz.True
    Assert.That List(3), Iz.False
End Sub

Public Sub Item_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    Dim List As BitArray
    Set List = Cor.NewBitArray(10)
    
    On Error Resume Next
    List(-1) = True
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Item_WithIndexGreaterThanItemCount_ThrowsArgumentOutOfRangeException()
    Dim List As BitArray
    Set List = Cor.NewBitArray(10)
    
    On Error Resume Next
    List(10) = True
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithValues_PerformsCopy(ByRef Target As Variant, ByRef Expected As Variant)
    AssertCopyTo Target, 0, Expected
End Sub

Public Sub CopyTo_WithValuesAndOffset_PerformatCopy(ByRef Target As Variant, ByRef Expected As Variant)
    AssertCopyTo Target, 1, Expected
End Sub

Private Sub AssertCopyTo(ByRef Target As Variant, ByVal Index As Long, ByRef Expected As Variant)
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(34)
    Bits(0) = True
    Bits(33) = True
    
    Bits.CopyTo Target, Index
    
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub CopyTo_WithUnsupportedArrayType_ThrowsArgumentException()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    On Error Resume Next
    Dim Target(0 To 1) As Integer
    Bits.CopyTo Target, 0
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub CopyTo_WithNullArray_ThrowsArgumentNullException()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    On Error Resume Next
    Dim NullTarget() As Byte
    Bits.CopyTo NullTarget, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub CopyTo_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    On Error Resume Next
    Dim Target(1 To 2) As Long
    Bits.CopyTo Target, 0
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithByteArrayTooSmall_ThrowsArgumentException()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    On Error Resume Next
    Dim Target(0) As Byte
    Bits.CopyTo Target, 0
    
    AssertArgumentException Err
End Sub

Public Sub CopyTo_WithBooleanArrayTooSmall_ThrowsArgumentException()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    On Error Resume Next
    Dim Target(0 To 8) As Boolean
    Bits.CopyTo Target, 0
    
    AssertArgumentException Err
End Sub

Public Sub FromBytes_WhenCalled_CreatesWithBytes()
    Dim Bits As BitArray
    
    Set Bits = BitArray.FromBytes(NewBytes(&HFF, 0, &HFF, 0, &HFF))

    Dim Actual(0 To 1) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HFF00FF, &HFF))
End Sub

Public Sub FromBytes_WithNullArray_ThrowsArgumentNullException()
    On Error Resume Next
    Dim NullBytes() As Byte
    
    BitArray.FromBytes NullBytes
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub FromLongs_WhenCalled_CreatesWithLongs()
    Dim Bits As BitArray
    
    Set Bits = BitArray.FromLongs(NewLongs(&HFF00FF00, &HFF))
    
    Dim Actual(0 To 1) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HFF00FF00, &HFF))
End Sub

Public Sub FromLongs_WithNullArray_ThrowsArgumentNullException()
    On Error Resume Next
    Dim NullLongs() As Long
    
    BitArray.FromLongs NullLongs
    
    AssertArgumentNullException Err, "Values"
End Sub


'Private list As bitarray
'
'Public Sub Setup()
'    Set list = newbitarray(32)
'End Sub
'
'Public Sub TestAnd()
'    Dim andlist As bitarray
'    Set andlist = newbitarray(list.Count)
'    Dim i As Long
'    For i = 0 To andlist.Count - 1
'        andlist(i) = Not CBool(i And 1)
'    Next i
'    list.SetAll True
'    list.AndBits andlist
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestOr()
'    Dim orlist As bitarray
'    Dim i As Long
'    Set orlist = newbitarray(list.Count)
'    For i = 0 To orlist.Count - 1
'        orlist(i) = Not CBool(i And 1)
'    Next i
'    list.OrBits orlist
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestSetAll()
'    Dim i As Long
'    list.SetAll True
'    For i = 0 To list.Count - 1
'        Assert.IsTrue list(i), "Failed at " & i
'    Next i
'
'    list.SetAll False
'    For i = 0 To list.Count - 1
'        Assert.IsFalse list(i), "Failed at " & i
'    Next i
'End Sub
'
'Public Sub TestSetLengthClearsNewBits()
'    Dim i As Long
'    i = list.Length
'    list.Length = list.Length + 35
'    list.SetAll True
'    list.Length = list.Length - 35
'    list.Length = list.Length + 35
'
'    For i = i To list.Count - 1
'        Assert.IsFalse list(i), "New bits were not cleared."
'    Next i
'End Sub
'
'Public Sub TestNot()
'    Dim i As Long
'
'    For i = 0 To list.Count - 1
'        list(i) = i And 1
'    Next i
'    list.NotBits
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestXor()
'    Dim i As Long
'    Dim b As bitarray
'    Set b = newbitarray(list.Count)
'    b.SetAll True
'    list.XorBits b
'
'    For i = 0 To list.Count - 1
'        Assert.IsTrue list(i), "Bit should be true."
'    Next i
'End Sub
'
'Public Sub TestFromLongs()
'    Dim l() As Long
'    l = NewLongs(&H55555555, &H55555555)
'
'    Set list = bitarray.FromLongs(l)
'    Assert.AreEqual 64, list.Count, "Count should be 64."
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestFromBytes()
'    Dim b() As Byte
'    b = NewBytes(&H55, &H55, &H55)
'
'    Set list = bitarray.FromBytes(b)
'    Assert.AreEqual 24, list.Count, "Count should be 24."
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestFromBooleans()
'    Dim b(60) As Boolean
'    Dim i As Long
'    For i = 0 To 60
'        b(i) = Not CBool(i And 1)
'    Next i
'    Set list = bitarray.FromBooleans(b)
'    Assert.AreEqual 61, list.Count, "Count should be 61."
'    AssertAlternatingPattern
'End Sub
'
'Public Sub TestEnumerator()
'    Dim v As Variant
'    list.SetAll True
'    For Each v In list
'        Assert.IsTrue v
'    Next v
'End Sub
'
