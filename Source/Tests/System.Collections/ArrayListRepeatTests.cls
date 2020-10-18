VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListRepeatTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ArrayListRepeatTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Public Sub Repeat_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    ArrayList.Repeat 1, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Repeat_WithCountOfZero_ReturnsEmptyList()
    Dim Actual As ArrayList
    
    Set Actual = ArrayList.Repeat(1, 0)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub Repeat_WithCount_ReturnsListContainingCountItems()
    Dim Actual As ArrayList
    
    Set Actual = ArrayList.Repeat("a", 10)
    
    Assert.That Actual.Count, Iz.EqualTo(10)
    Assert.That Actual, Has.All.EqualTo("a")
End Sub
