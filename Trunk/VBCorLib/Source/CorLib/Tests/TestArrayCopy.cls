VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayCopyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestArrayCopy
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Array Tests", "Array Copy Tests"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Copy_WithValues_CopiesWithConversion"
            Test.Use NewBytes(1, 2, 3), NewBytes(0, 0, 0), NewBytes(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewIntegers(0, 0, 0), NewIntegers(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewLongs(0, 0, 0), NewLongs(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewSingles(0, 0, 0), NewSingles(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewCurrencys(0, 0, 0), NewCurrencys(1, 2, 3)
            Test.Use NewBytes(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewBytes(1, 2, 3), NewVariants(0, 0, 0), NewVariants(CByte(1), CByte(2), CByte(3))
            
            Test.Use NewIntegers(1, 2, 3), NewIntegers(0, 0, 0), NewIntegers(1, 2, 3)
            Test.Use NewIntegers(1, 2, 3), NewLongs(0, 0, 0), NewLongs(1, 2, 3)
            Test.Use NewIntegers(1, 2, 3), NewSingles(0, 0, 0), NewSingles(1, 2, 3)
            Test.Use NewIntegers(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            Test.Use NewIntegers(1, 2, 3), NewCurrencys(0, 0, 0), NewCurrencys(1, 2, 3)
            Test.Use NewIntegers(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewIntegers(1, 2, 3), NewVariants(0, 0, 0), NewVariants(1, 2, 3)
            
            Test.Use NewLongs(1, 2, 3), NewLongs(0, 0, 0), NewLongs(1, 2, 3)
            Test.Use NewLongs(1, 2, 3), NewSingles(0, 0, 0), NewSingles(1, 2, 3)
            Test.Use NewLongs(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            Test.Use NewLongs(1, 2, 3), NewCurrencys(0, 0, 0), NewCurrencys(1, 2, 3)
            Test.Use NewLongs(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewLongs(1, 2, 3), NewVariants(0, 0, 0), NewVariants(1&, 2&, 3&)
            
            Test.Use NewSingles(1, 2, 3), NewSingles(0, 0, 0), NewSingles(1, 2, 3)
            Test.Use NewSingles(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            Test.Use NewSingles(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewSingles(1, 2, 3), NewVariants(0, 0, 0), NewVariants(1!, 2!, 3!)
            
            Test.Use NewDoubles(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            Test.Use NewDoubles(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewDoubles(1, 2, 3), NewVariants(0, 0, 0), NewVariants(1#, 2#, 3#)
            
            Test.Use NewCurrencys(1, 2, 3), NewCurrencys(0, 0, 0), NewCurrencys(1, 2, 3)
            Test.Use NewCurrencys(1, 2, 3), NewStrings("", "", ""), NewStrings("1", "2", "3")
            Test.Use NewCurrencys(1, 2, 3), NewVariants(0, 0, 0), NewVariants(1@, 2@, 3@)
            Test.Use NewCurrencys(1, 2, 3), NewDoubles(0, 0, 0), NewDoubles(1, 2, 3)
            
            Dim Objects(0 To 2) As Int32
            Set Objects(0) = New Int32
            Set Objects(1) = New Int32
            Set Objects(2) = New Int32
            Test.Use Objects, Array(Nothing, Nothing, Nothing), Objects
    End Select
End Sub

Public Sub Copy_WithNegativeLength_Throws()
    On Error GoTo errTrap
    Dim Arr(0) As Long
    
    cArray.Copy Arr, Arr, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub Copy_WithValues_CopiesWithConversion(ByRef Source As Variant, ByRef Destination As Variant, ByRef Expected As Variant)
    cArray.Copy Source, Destination, 3
    
    Assert.That Destination, Iz.EqualTo(Expected).Strict
End Sub

Public Sub Copy_WithLengthGreaterThanSourceLength_Throws()
    On Error GoTo errTrap
    Dim Arr1(0) As Long
    Dim Arr2(2) As Long
    
    cArray.Copy Arr1, Arr2, 2
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub Copy_WithLengthGreaterThanDestinationLength_Throws()
    On Error GoTo errTrap
    Dim Arr1(0) As Long
    Dim Arr2(2) As Long
    
    cArray.Copy Arr2, Arr1, 2
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub CopyEx_WithLengthPlusStartIndexGreaterThanSourceLength_Throws()
    On Error GoTo errTrap
    Dim Arr1(1) As Long
    Dim Arr2(2) As Long
    
    cArray.Copyex Arr1, 1, Arr2, 0, 2
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub CopyEx_WithLengthPlusDestinationIndexGreaterThanDestinationLength_Throws()
    On Error GoTo errTrap
    Dim Arr1(1) As Long
    Dim Arr2(2) As Long
    
    cArray.Copyex Arr2, 0, Arr1, 1, 2
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub CopyEx_WithSourceIndexLessThanLBound_Throws()
    On Error GoTo errTrap
    Dim Arr1(1 To 1) As Long
    Dim Arr2(1 To 1) As Long
    
    cArray.Copyex Arr1, 0, Arr2, 1, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "SourceIndex"
End Sub

Public Sub CopyEx_WithDestinationIndexLessThanLBound_Throws()
    On Error GoTo errTrap
    Dim Arr1(1 To 1) As Long
    Dim Arr2(1 To 1) As Long
    
    cArray.Copyex Arr1, 1, Arr2, 0, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "DestinationIndex"
End Sub

Public Sub Copy_WithNullSource_Throws()
    On Error GoTo errTrap
    Dim NullArray() As Long
    Dim NonNull(0) As Long
    
    cArray.Copy NullArray, NonNull, 1
    
errTrap:
    AssertArgumentNullException Err, "SourceArray"
End Sub

Public Sub Copy_WithNullDestination_Throws()
    On Error GoTo errTrap
    Dim NullArray() As Long
    Dim NonNull(0) As Long
    
    cArray.Copy NonNull, NullArray, 1
    
errTrap:
    AssertArgumentNullException Err, "DestinationArray"
End Sub

Public Sub CopyEx_WithNegativeLength_Throws()
    On Error GoTo errTrap
    Dim Arr(0) As Long
    
    cArray.Copyex Arr, 0, Arr, 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub CopyEx_WithNullSource_Throws()
    On Error GoTo errTrap
    Dim NullArray() As Long
    Dim NonNull(0) As Long
    
    cArray.Copyex NullArray, 0, NonNull, 0, 1
    
errTrap:
    AssertArgumentNullException Err, "SourceArray"
End Sub

Public Sub CopyEx_WithNullDestination_Throws()
    On Error GoTo errTrap
    Dim NullArray() As Long
    Dim NonNull(0) As Long
    
    cArray.Copyex NonNull, 0, NullArray, 0, 1
    
errTrap:
    AssertArgumentNullException Err, "DestinationArray"
End Sub

Public Sub Copy_WithLongs_CopiesAllElements()
    Dim Source() As Long
    Dim Target() As Long
    Source = NewLongs(1, 2, 3, 4, 5)
    ReDim Target(0 To 4)
    
    cArray.Copy Source, Target, 5
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub Copy_WithStrings_CopiesAllElements()
    Dim Source() As String
    Dim Target() As String
    Source = NewStrings("1", "2", "3", "4", "5")
    ReDim Target(0 To 4)
    
    cArray.Copy Source, Target, 5
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub Copy_WithCountLessThanArrayLength_CopiesSpecifiedNumberOfElements()
    Dim Source() As Long
    Dim Target() As Long
    Source = NewLongs(1, 2, 3, 4, 5)
    ReDim Target(0 To 4)
    
    cArray.Copy Source, Target, 3
    
    Dim Expected() As Long
    Expected = NewLongs(1, 2, 3, 0, 0)
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithUnalignedLBounds_CopiesElements()
    Dim Source() As Long
    Dim Target() As Long
    Source = NewLongs(1, 2, 3, 4, 5)
    ReDim Target(5 To 9)
    
    cArray.Copy Source, Target, 5
    
    Dim Expected(5 To 9) As Long
    Expected(5) = 1
    Expected(6) = 2
    Expected(7) = 3
    Expected(8) = 4
    Expected(9) = 5
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithCallback_CopiesElementsUsingCallback()
    Dim Source() As String
    Dim Target() As String
    Source = NewStrings("11", "22", "33")
    ReDim Target(0 To 2)
    
    cArray.Copy Source, Target, 3, AddressOf modTestCallbacks.FirstLetterCopier
    
    Dim Expected() As String
    Expected = NewStrings("1", "2", "3")
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_ToSelfUsingCallback_CopiesElementsUsingCallback()
    Dim Source() As String
    Source = NewStrings("11", "22", "33")
    
    cArray.Copy Source, Source, 3, AddressOf modTestCallbacks.FirstLetterCopier
    
    Dim Expected() As String
    Expected = NewStrings("1", "2", "3")
    Assert.That Source, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithMultiDimensionalArrays_CopiesElements()
    Dim Source(0 To 2, 0 To 2) As Long
    Dim Target(0 To 2, 0 To 2) As Long
    Source(0, 0) = 1
    Source(1, 1) = 2
    Source(2, 2) = 3

    cArray.Copy Source, Target, 9

    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub CopyEx_WithLongsSubSet_CopiesSubSet()
    Dim Source() As Long
    Dim Target() As Long
    Source = NewLongs(1, 2, 3, 4, 5, 6, 7, 8)
    Target = NewLongs(0, 0, 0, 0, 0, 0, 0, 0)
    
    cArray.Copyex Source, 2, Target, 0, 4
    
    Dim Expected() As Long
    Expected = NewLongs(3, 4, 5, 6, 0, 0, 0, 0)
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithObjects_CopiesElements()
    Dim Source(1) As Int32
    Dim Target(1) As Int32
    Set Source(0) = New Int32
    Set Source(1) = New Int32
    
    cArray.Copy Source, Target, 2
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub CopyEx_WithObjects_CopiesElements()
    Dim Source(0 To 1) As Int32
    Dim Target(1 To 2) As Int32
    Set Source(0) = New Int32
    Set Source(1) = New Int32
    
    cArray.Copyex Source, 1, Target, 2, 1
    
    Assert.That Target(1), Iz.Nothing
    Assert.That Target(2), Iz.SameAs(Source(1))
End Sub

Public Sub TestObjectsToSelf()
    Dim Item1 As New Int32
    Dim Item2 As New Int32
    Dim Item3 As New Int32
    Dim actual(0 To 3) As Int32
    Set actual(0) = Item1
    Set actual(1) = Item2
    Set actual(2) = Item3
    Set actual(3) = New Int32
    
    cArray.Copyex actual, 0, actual, 1, 3
    
    Dim Expected(0 To 3) As Int32
    Set Expected(0) = Item1
    Set Expected(1) = Item1
    Set Expected(2) = Item2
    Set Expected(3) = Item3
    Assert.That actual, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithLongsToStrings_ConvertsElements()
    Dim Source() As Long
    Dim Target() As String
    Source = NewLongs(1, 2, 3)
    ReDim Target(2)
    
    cArray.Copy Source, Target, 3
    
    Dim Expected() As String
    Expected = NewStrings("1", "2", "3")
    Assert.That Target, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithVariantsToObjects_ConvertsElements()
    Dim Source As Variant
    Dim Target() As Int32
    Source = Array(New Int32, New Int32, New Int32)
    ReDim Target(2)
    
    cArray.Copy Source, Target, 3
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub Copy_WithIncompatibleTypes_Throws()
    Dim Source(0) As Double
    Dim Target(0) As Long
    
    On Error GoTo errTrap
    cArray.Copy Source, Target, 1
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArrayTypeMismatchException, "Wrong exception type caught."
End Sub

Public Sub TestCopyToSelf()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i + 1
    Next i
    
    cArray.Copyex a, 0, a, 1, 100
    
    Assert.AreEqual 1, a(0)
    
    For i = 1 To 100
        Assert.AreEqual i, a(i)
    Next i
End Sub

Public Sub TestUserDefined()
    Dim src(10) As VBGUID
    Dim dst(10) As VBGUID
    Dim i As Long

    For i = 0 To 10
        src(i).Data1 = i
    Next i

    cArray.Copy src, dst, 11

    For i = 0 To 10
        Assert.AreEqual i, dst(i).Data1
    Next i

End Sub

Public Sub TestCopyObjectsWithCast()
    Dim src(10) As Int32
    Dim dst(10) As IComparable
    Dim i As Long
    
    For i = 0 To 10
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    cArray.Copy src, dst, 11
    
    For i = 0 To 10
        Assert.IsNotNothing dst(i)
        Assert.AreEqual 0, dst(i).CompareTo(src(i))
    Next i
    
End Sub

Public Sub TestCopyObjectsMismatch()
    Dim src(10) As Int32
    Dim dst(10) As Int32Formattable
    Dim i As Long
    
    For i = 0 To 10
        Set src(10) = New Int32
    Next i
    
    On Error GoTo errTrap
    cArray.Copy src, dst, 11
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArrayTypeMismatchException, "Wrong exception type caught."
End Sub

Public Sub TestNonZeroLBound()
    Dim src(1 To 10) As String
    Dim dst(1 To 10) As String
    Dim i As Long
    
    For i = 1 To 10
        src(i) = "Item" & i
    Next i
    
    cArray.Copy src, dst, 10
    
    For i = 1 To 10
        Assert.AreEqual "Item" & i, dst(i)
    Next i
End Sub



