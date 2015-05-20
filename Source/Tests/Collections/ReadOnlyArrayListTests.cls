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

Public Sub Add_WhenCalled_ThrowsInvalidOperationException()
    On Error Resume Next
    
    With MakeList
        .Add 1
    End With
    
    Assert
End Sub


Private Function MakeList() As ArrayList
    Set MakeList = ArrayList.ReadOnly(Cor.NewArrayList(Source:=Array(1, 2, 3)))
End Function
