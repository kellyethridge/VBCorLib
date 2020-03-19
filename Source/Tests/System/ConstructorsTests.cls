VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstructorsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConstructorsTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub NewDecimals_WithInt64_ReturnsDecimalArray()
    Dim Actual() As Variant
    
    Actual = NewDecimals(CInt64(1))

    Assert.That Actual, Iz.EqualTo(Array(CDec(1))).Strict
End Sub

Public Sub NewInt64s_WhenCalled_ReturnsInt64Array()
    Dim Actual() As Int64
    
    Actual = NewInt64s(1, 2, 3)
    
    Assert.That Actual, Iz.EqualTo(Array(CInt64(1), CInt64(2), CInt64(3)))
End Sub

