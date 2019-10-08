VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharEnumeratorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CharEnumeratorTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub MoveNext_WithEmptyString_ReturnsFalse()
    Dim Enumerator As IEnumerator
    Set Enumerator = Cor.NewCharEnumerator("")
    
    Dim Actual As Boolean
    Actual = Enumerator.MoveNext
    
    Assert.That Actual, Iz.False
End Sub

Public Sub MoveNext_WhenCalled_MovesToNextCharactor()
    Dim Enumerator As CharEnumerator
    Set Enumerator = Cor.NewCharEnumerator("hello")
    
    Dim Count       As Long
    Dim Actual      As Integer
    Dim Expected    As Integer
    Do While Enumerator.MoveNext
        Count = Count + 1
        Expected = Choose(Count, 104, 101, 108, 108, 111)
        
        Actual = Enumerator.Current
        
        Assert.That Actual, Iz.EqualTo(Expected)
    Loop
End Sub

Public Sub MoveNext_AfterLastCharacter_ReturnsFalse()
    Dim Enumerator As CharEnumerator
    Set Enumerator = Cor.NewCharEnumerator("a")
    Enumerator.MoveNext
    
    Dim Actual As Boolean
    Actual = Enumerator.MoveNext
    
    Assert.That Actual, Iz.False
End Sub

Public Sub MoveNext_AfterFinalMoveNext_ThrowsInvalidOperationException()
    Dim Enumerator As CharEnumerator
    Set Enumerator = Cor.NewCharEnumerator("a")
    Enumerator.MoveNext
    Enumerator.MoveNext

    On Error Resume Next
    Enumerator.MoveNext
    
    AssertInvalidOperationException Err
End Sub

Public Sub Current_AfterLastCharactor_ThrowsInvalidOperationException()
    Dim Enumerator As CharEnumerator
    Set Enumerator = Cor.NewCharEnumerator("a")
    Enumerator.MoveNext
    Enumerator.MoveNext

    On Error Resume Next
    Dim Value As Variant
    Value = Enumerator.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub Current_BeforeStartingEnumeration_ThrowsInvalidOperationException()
    Dim Enumerator As CharEnumerator
    Set Enumerator = Cor.NewCharEnumerator("a")
    
    On Error Resume Next
    Dim Value As Variant
    Value = Enumerator.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub ForEach_WhenCalled_EnumeratesCharacters()
    Dim s As String
    Dim sb As New StringBuilder
    Dim Ch As Variant
    
    s = "Hello"
    For Each Ch In Cor.NewCharEnumerator(s)
        sb.AppendChar Ch
    Next
    
    Assert.That sb.ToString, Iz.EqualTo(s)
End Sub
