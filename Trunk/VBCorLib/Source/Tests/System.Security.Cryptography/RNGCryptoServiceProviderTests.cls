VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RNGCryptoServiceProviderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RNGCryptoServiceProviderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub GetBytes_WhenCalled_FilledArray()
    Dim Generator As New RNGCryptoServiceProvider
    
    Dim Actual(1 To 100) As Byte
    Generator.GetBytes Actual
    
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub GetBytes_WhenCalled_RandomlyFillsArrays()
    Dim Generator As New RNGCryptoServiceProvider
    
    Dim Actual1(1 To 10) As Byte
    Dim Actual2(1 To 10) As Byte
    Generator.GetBytes Actual1
    Generator.GetBytes Actual2
    
    Assert.That Actual1, Iz.Not.EqualTo(Actual2)
End Sub

Public Sub GetNonZeroBytes_WhenCalled_FillArrayWithNonZeroRandomValues()
    Dim Generator As New RNGCryptoServiceProvider
    
    Dim Actual(1 To 100) As Byte
    Generator.GetNonZeroBytes Actual
    
    Assert.That Actual, Has.None.EqualTo(0)
End Sub

Public Sub GetNonZeroBytes_WhenCalled_RandomlyFillsArrays()
    Dim Generator As New RNGCryptoServiceProvider
    
    Dim Actual1(1 To 10) As Byte
    Dim Actual2(1 To 10) As Byte
    Generator.GetNonZeroBytes Actual1
    Generator.GetNonZeroBytes Actual2
    
    Assert.That Actual1, Iz.Not.EqualTo(Actual2)
End Sub
