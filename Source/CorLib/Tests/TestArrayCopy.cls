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

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Array Tests", "Array Copy Tests"
End Sub

Public Sub Copy_WithLongs_CopiesAllElements()
    Dim Source() As Long
    Dim Target() As Long
    Source = NewLongs(1, 2, 3, 4, 5)
    ReDim Target(0 To 4)
    
    cArray.Copy Source, Target, 5
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub Copyt_WithStrings_CopiesAllElements()
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

Public Sub Copy_WithMultiDimensionalArrays_CopiesAllElements()
    Assert.Ignore "Not implemented."
    Dim Source(0 To 2, 0 To 2) As Long
    Dim Target(0 To 2, 0 To 2) As Long
    Source(0, 0) = 1
    Source(1, 1) = 2
    Source(2, 2) = 3
    
    cArray.Copy Source, Target, 9
    
    Assert.That Target, Iz.EqualTo(Source)
End Sub

Public Sub CopyEx_WithLongsSubSet_CopiesSubSet()
    Dim src() As Long
    Dim dst() As Long
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    cArray.CopyEx src, 20, dst, 0, 20
    For i = 0 To 100
        If i >= 20 Then
            Assert.AreEqual 0, dst(i)
        Else
            Assert.AreEqual i + 20, dst(i)
        End If
    Next i

End Sub

Public Sub TestCopyObjects()
    Dim src() As Int32
    Dim dst() As Int32
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    cArray.Copy src, dst, 101
    For i = 0 To 100
        Assert.AreEqual i, dst(i).mValue
    Next i
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
    
    cArray.CopyEx actual, 0, actual, 1, 3
    
    Dim Expected(0 To 3) As Int32
    Set Expected(0) = Item1
    Set Expected(1) = Item1
    Set Expected(2) = Item2
    Set Expected(3) = Item3
    Assert.That actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestLongToString()
    Dim src() As Long
    Dim dst() As String
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    cArray.Copy src, dst, 101
    For i = 0 To 100
        Assert.AreEqual CStr(i), dst(i)
    Next i
End Sub

Public Sub TestIndexes()
    Dim src() As String
    Dim dst() As String
    
    ReDim src(0 To 4)
    ReDim dst(5 To 15)
    
    Dim i As Long
    For i = 0 To 4
        src(i) = i
    Next i
    
    cArray.CopyEx src, 0, dst, 5, 5
    
    For i = 5 To 9
        Assert.AreEqual CStr(i - 5), dst(i)
    Next i
End Sub

Public Sub TestVariantToObjects()
    Dim src() As Variant
    Dim dst() As Int32
    
    ReDim src(0 To 5)
    ReDim dst(0 To 5)
    Dim i As Long
    For i = 0 To 5
        Set src(i) = New Int32
        Set dst(i) = New Int32
    Next i
    
    cArray.Copy src, dst, 6
    
    For i = 0 To 5
        Assert.AreEqual src(i), dst(i)
    Next i
End Sub

Public Sub TestArrayTypeMismatch()
    Dim src(0) As Double
    Dim dst(0) As Long
    
    On Error GoTo errTrap
    cArray.Copy src, dst, 1
    On Error GoTo 0
    Assert.Fail "Error should have been raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArrayTypeMismatchException, "Wrong exception type caught."
    
End Sub

Public Sub TestCopyToSelf()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i + 1
    Next i
    
    cArray.CopyEx a, 0, a, 1, 100
    
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
    Assert.IsTrue Catch(Ex), "An exception should be caught."
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


