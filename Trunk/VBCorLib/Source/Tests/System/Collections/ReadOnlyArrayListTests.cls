VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyArrayListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Public Sub ReadOnly_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    ArrayList.ReadOnly Nothing
    
    AssertArgumentNullException Err, "List"
End Sub

Public Sub IsReadOnly_WhenCalled_ReturnsTrue()
    Dim Actual As Boolean
    
    With MakeList
        Actual = .IsReadOnly
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

Public Sub Item_WhenCalled_ReturnsValue()
    Dim Actual As Long
    
    With MakeList
        Actual = .Item(1)
    End With
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub Item_WhenSet_ThrowsNotSupportedException()
    On Error Resume Next
    
    With MakeList
        .Item(1) = 5
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

Public Sub Clone_WhenCalled_ReturnsReadOnlyArrayList()
    Dim Clone As ArrayList
    Dim Actual As Boolean
    
    With MakeList
        Set Clone = .Clone
        Actual = Clone.IsReadOnly
    End With
    
    Assert.That Actual, Iz.True
End Sub

Private Function MakeList() As ArrayList
    Set MakeList = ArrayList.ReadOnly(Cor.NewArrayList(Source:=Array(1, 2, 3)))
End Function
