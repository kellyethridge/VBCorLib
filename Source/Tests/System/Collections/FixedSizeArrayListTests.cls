VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixedSizeArrayListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FixedSize_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    ArrayList.FixedSize Nothing
    
    AssertArgumentNullException Err, "List"
End Sub

Public Sub IsFixedSize_WhenCalled_ReturnsTrue()
    Dim Actual As Boolean
    
    With MakeList
        Actual = .IsFixedSize
    End With
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Add_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .Add 1
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub Remove_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .Remove 1
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub RemoveAt_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .RemoveAt 1
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub RemoveRange_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .RemoveRange 1, 1
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub Insert_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .Insert 0, 4
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub AddRange_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .AddRange Array(1)
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub Clear_WhenCalled_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .Clear
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub Clone_WhenCalled_ReturnsFixedSizeArrayList()
    Dim Clone As ArrayList
    Dim Actual As Boolean
    
    With MakeList
        Set Clone = .Clone
        Actual = Clone.IsFixedSize
    End With
    
    Assert.That Actual, Iz.True
End Sub

Public Sub BinarySearch_WhenCalled_ReturnsExpected()
    Dim Actual As Long
    
    With MakeList
        Actual = .BinarySearch(3)
    End With
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetEnumerator_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeList
    On Error GoTo Catch
    
    List.GetEnumerator -1, 1
    
Catch:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetEnumerator_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeList
    On Error GoTo Catch
    
    List.GetEnumerator 0, -1
    
Catch:
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetEnumerator_WithIndexPlusCountGreaterThanCount_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeList
    On Error GoTo Catch
    
    List.GetEnumerator 1, 10
    
Catch:
    AssertArgumentException Err
End Sub

Public Sub ForEach_WhenCalled_EnumeratesList()
    Dim List As ArrayList
    Set List = MakeList
    Dim Item As Variant
    Dim Actual As Long
    
    For Each Item In List
        Actual = Actual + 1
    Next
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub



Private Function MakeList() As ArrayList
    Set MakeList = ArrayList.FixedSize(Cor.NewArrayList(Source:=Array(1, 2, 3)))
End Function
