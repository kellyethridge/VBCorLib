VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdapterBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdapterBinarySearch
'
Option Explicit

Public Sub TestSearchSubArray()
    Dim a As ArrayList
    Dim i As Long
    
    Set a = New ArrayList
    For i = 0 To 100
        Select Case i
            Case 26 To 49
                a.Add i
            Case Else
                a.Add 0
        End Select
    Next i
    
    Set a = ArrayList.Adapter(a)
    
    For i = 26 To 49
        Assert.AreEqual i, a.BinarySearch(i, 26, 24)
    Next i
End Sub

Public Sub TestNotInSubArray()
    Dim a As ArrayList
    Dim i As Long

    Set a = New ArrayList
    For i = 0 To 100
        Select Case i
            Case 26 To 49
                a.Add i
            Case Else
                a.Add 0
        End Select
    Next i

    Set a = ArrayList.Adapter(a)

    For i = 0 To 25
        Assert.AreEqual Not 26, a.BinarySearch(i, 26, 24, comparer.Default)
    Next i
    For i = 50 To 100
        Assert.AreEqual Not 50, a.BinarySearch(i, 26, 24)
    Next i
End Sub

Public Sub TestObjectBinarySearch()
    Dim a As ArrayList
    Dim i As Long

    Set a = New ArrayList
    For i = 0 To 100
        Dim obj As Int32
        Set obj = New Int32
        obj.mValue = i
        a.Add obj
    Next i
    
    Set a = ArrayList.Adapter(a)
    
    Dim oInt As New Int32
    oInt.mValue = 5
    i = a.BinarySearch(oInt)
    Assert.AreEqual 5, i, "Should return 5."
End Sub

Public Sub TestStringBinarySearch()
    Dim a As ArrayList
    Dim i As Long

    Set a = New ArrayList
    For i = 0 To 100
        a.Add "Item " & Chr$(i)
    Next i

    Set a = ArrayList.Adapter(a)

    i = a.BinarySearch("Item " & Chr$(5))
    Assert.AreEqual 5, i, "Should return 5."
End Sub

Public Sub TestVariantBinarySearch()
    Dim a As ArrayList
    Dim i As Long

    Set a = New ArrayList
    For i = 0 To 100
        a.Add "Item " & Chr$(i)
    Next i
    
    Set a = ArrayList.Adapter(a)
    
    a(0) = Null
    a(1) = Empty
    Assert.AreEqual 0, a.BinarySearch(Null), "Should return 0."
    Assert.AreEqual 1, a.BinarySearch(Empty), "Should return 1."
End Sub

Public Sub TestLengthTooLong()
    Dim a As ArrayList
    Set a = New ArrayList
    a.Capacity = 10
    
    Set a = ArrayList.Adapter(a)

    On Error GoTo errTrap
    a.BinarySearch 1, 0, 12
    On Error GoTo 0
    Assert.Fail "An exception should have been raised."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

