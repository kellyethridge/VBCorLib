VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBRndTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBRndTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub TwoDifferentNumbers()
    Dim b1 As BigInteger
    Dim b2 As BigInteger
    
    Set b1 = BigInteger.Rnd(20)
    Set b2 = BigInteger.Rnd(20)
    
    Assert.IsNotNothing b1
    Assert.IsNotNothing b2
    
    Assert.AreEqual 1, b1.Sign
    Assert.AreEqual 1, b2.Sign
    Assert.AreEqual 20, b1.Precision
    Assert.AreEqual 20, b2.Precision
    
    Assert.IsFalse b1.Equals(b2)
    Assert.IsFalse b1.Equals(BigInteger.Zero)
    Assert.IsFalse b2.Equals(BigInteger.Zero)
End Sub

Public Sub SameSeed()
    Dim b1 As BigInteger
    Dim b2 As BigInteger
    
    BigInteger.SeedRnd -13
    Set b1 = BigInteger.Rnd(20)
    
    BigInteger.SeedRnd -13
    Set b2 = BigInteger.Rnd(20)
    
    Assert.IsTrue b1.Equals(b2)
End Sub

Public Sub GenerateNegative()
    Dim b1 As BigInteger
    Dim b2 As BigInteger
    
    BigInteger.SeedRnd -13
    Set b1 = BigInteger.Rnd(20, True)
    
    Assert.AreEqual -1, b1.Sign
    
    BigInteger.SeedRnd -13
    Set b2 = BigInteger.Rnd(20)
        
    Assert.AreEqual 1, b2.Sign
    
    Assert.IsFalse b1.Equals(b2)
End Sub
