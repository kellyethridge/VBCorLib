VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBFactorialTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBFactorialTests
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

Public Sub Factorial0()
    Set b = BigInteger.Factorial(BInt(0))
    Assert.AreEqual Sign.Zero, b.Sign
End Sub

Public Sub Factorial1()
    Set b = BigInteger.Factorial(BInt(1))
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub Factorial2()
    Set b = BigInteger.Factorial(BInt(2))
    Assert.AreEqual "2", b.ToString
End Sub

Public Sub Factorial3()
    Set b = BigInteger.Factorial(BInt(3))
    Assert.AreEqual "6", b.ToString
End Sub

Public Sub Factorial6()
    Set b = BigInteger.Factorial(BInt(6))
    Assert.AreEqual "720", b.ToString
End Sub

Public Sub Factorial20()
    Set b = BigInteger.Factorial(BInt(20))
    Assert.AreEqual "2432902008176640000", b.ToString
End Sub
