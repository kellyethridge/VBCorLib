VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBSquareRootTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBSquareRootTests
'
Option Explicit
Implements ICategorizable

Private b As BigInteger

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub Test4()
    AssertSqrt 4, 2
End Sub

Public Sub Test8()
    AssertSqrt 8, 2
End Sub

Public Sub Test9()
    AssertSqrt 9, 3
End Sub

Public Sub Test16()
    AssertSqrt 16, 4
End Sub

Public Sub Test15()
    AssertSqrt 15, 3
End Sub

Public Sub Test81()
    AssertSqrt 81, 9
End Sub

Public Sub Test1000()
    AssertSqrt 1000, 31
End Sub

Public Sub Test1000000()
    AssertSqrt 1000000, 1000
End Sub

Public Sub Test999999()
    AssertSqrt 999999, 999
End Sub

Public Sub TestLarge()
    AssertSqrt "123456789012345678901234567", "11111111061111"
End Sub

Private Sub AssertSqrt(ByVal Value As Variant, ByVal Expected As String)
    Set b = BInt(Value)
    Set b = BigInteger.SquareRoot(b)
    
    Assert.IsNotNothing b, "Should not be Nothing"
    Assert.AreEqual 1, b.Sign, "Wrong sign"
    Assert.AreEqual Expected, b.ToString
End Sub
