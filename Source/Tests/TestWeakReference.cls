VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestWeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestWeakReference
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub TestWeakness()
    Dim Strong As StrongRef
    Dim Weak As WeakRef
    
    Set Strong = New StrongRef
    Set Weak = New WeakRef
    Set Strong.Reference = Weak
    Set Weak.Reference = Strong
    
    Assert.IsNotNothing Strong.Reference, "Target should not be Nothing."
    
    UseStrongRef Strong
    
    Strong.AssertAlive
    Set Weak = Nothing
    Strong.AssertDead
End Sub

Private Sub UseStrongRef(ByVal Ref As StrongRef)
    Ref.AssertAlive
End Sub
