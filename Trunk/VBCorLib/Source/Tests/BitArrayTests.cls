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
            Test.Use NewLongs(0, 0), NewLongs(&HFFFFFFFF, &H3)
            Test.Use NewBytes(0, 0, 0, 0, 0), NewBytes(&HFF, &HFF, &HFF, &HFF, &H3)
            Test.Use CorArray.CreateInstance(vbBoolean, 34), NewBooleans(True, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True)
            
        Case "CopyTo_WithValuesAndOffset_PerformsCopy"
            Test.Use NewLongs(0, 0, 0), NewLongs(0, &HFFFFFFFF, &H3)
            Test.Use NewBytes(0, 0, 0, 0, 0, 0), NewBytes(0, &HFF, &HFF, &HFF, &HFF, &H3)
            Test.Use CorArray.CreateInstance(vbBoolean, 35), NewBooleans(False, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True, True, True, True, True, True, True, _
                                                                         True, True, True)
        
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

Public Sub Length_WithValueLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    With Cor.NewBitArray(10)
        .Length = -1
    End With
    
    AssertArgumentOutOfRangeException Err, "Value"
End Sub

Public Sub Length_WhenShortened_TruncatesElements()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(34)
    Bits(33) = True
    
    Bits.Length = 33
    
    Dim Actual(0 To 1) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(0, 0))
End Sub

Public Sub Length_WhenLengthend_CreatesNewElements()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(16, True)
    
    Bits.Length = 32
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HFFFF&))
End Sub

Public Sub CopyTo_WithValues_PerformsCopy(ByRef Target As Variant, ByRef Expected As Variant)
    AssertCopyTo Target, 0, Expected
End Sub

Public Sub CopyTo_WithValuesAndOffset_PerformsCopy(ByRef Target As Variant, ByRef Expected As Variant)
    AssertCopyTo Target, 1, Expected
End Sub

Private Sub AssertCopyTo(ByRef Target As Variant, ByVal Index As Long, ByRef Expected As Variant)
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(34, True)
    
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

Public Sub SetAll_WithTrue_SetsBitsToOne()
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    Bits.SetAll True
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&H3FF))
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

Public Sub FromBooleans_WhenCalled_CreatesWithBooleans()
    Dim Bits As BitArray
    
    Set Bits = BitArray.FromBooleans(NewBooleans(True, False, True, False))
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&H5))
End Sub

Public Sub FromBooleans_WithNullArray_ThrowsArgumentNullException()
    On Error Resume Next
    Dim NullBooleans() As Boolean
    
    BitArray.FromBooleans NullBooleans
    
    AssertArgumentNullException Err, "Values"
End Sub

Public Sub AndBits_WithValue_PerformsAndOperation()
    Dim Bits As BitArray
    Set Bits = BitArray.FromLongs(NewLongs(&HFFFFFFFF))
    
    Bits.AndBits CreateMask
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&H55555555))
End Sub

Public Sub AndBits_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    Bits.AndBits Nothing
    
    AssertArgumentNullException Err, "Bits"
End Sub

Public Sub AndBits_WithDifferentNumberOfElements_ThrowsArgumentException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(31)
    
    Bits.AndBits CreateMask
    
    AssertArgumentException Err
End Sub

Public Sub OrBits_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    Bits.OrBits Nothing
    
    AssertArgumentNullException Err, "Bits"
End Sub

Public Sub OrBits_WithValue_PerformsOrOperation()
    Dim Bits As BitArray
    Set Bits = BitArray.FromLongs(NewLongs(&H80000000))
    
    Bits.OrBits CreateMask
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HD5555555))
End Sub

Public Sub OrBits_WithDifferentNumberOfElements_ThrowsArgumentException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(31)
    
    Bits.OrBits CreateMask
    
    AssertArgumentException Err
End Sub

Public Sub XorBits_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(10)
    
    Bits.XorBits Nothing
    
    AssertArgumentNullException Err, "Bits"
End Sub

Public Sub XorBits_WithValue_PerformsOrOperation()
    Dim Bits As BitArray
    Set Bits = BitArray.FromLongs(NewLongs(&HFFFFFFFF))
    
    Bits.XorBits CreateMask
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HAAAAAAAA))
End Sub

Public Sub XorBits_WithDifferentNumberOfElements_ThrowsArgumentException()
    On Error Resume Next
    Dim Bits As BitArray
    Set Bits = Cor.NewBitArray(31)
    
    Bits.XorBits CreateMask
    
    AssertArgumentException Err
End Sub

Public Sub NotBits_WithValue_PerformsOrOperation()
    Dim Bits As BitArray
    Set Bits = BitArray.FromLongs(NewLongs(&H55555555))
    
    Bits.NotBits
    
    Dim Actual(0) As Long
    Bits.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&HAAAAAAAA))
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim Bits As BitArray
    Set Bits = CreateMask
    
    Dim Bits2 As BitArray
    Set Bits2 = Persist(Bits)
    
    Dim Actual(0) As Long
    Bits2.CopyTo Actual, 0
    Assert.That Actual, Iz.EqualTo(NewLongs(&H55555555))
End Sub

Public Sub New_WhenCalled_InitializesTo32Bits()
    Dim Bits As New BitArray
    
    Bits(31) = True
    
    Assert.That Bits.Length, Iz.EqualTo(32)
    Assert.That Bits(31), Iz.True
End Sub


Private Function CreateMask() As BitArray
    Set CreateMask = BitArray.FromLongs(NewLongs(&H55555555))
End Function

'
