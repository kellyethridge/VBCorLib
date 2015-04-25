VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoStream
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub TestCreateReadable()
    Dim ms As New MemoryStream
    Dim tr As New FromBase64Transform
    Dim cs As CryptoStream
    Set cs = Cor.NewCryptoStream(ms, tr, ReadMode)
    Assert.IsTrue cs.CanRead
    Assert.IsFalse cs.CanWrite
End Sub

Public Sub TestCreateWritable()
    Dim ms As New MemoryStream
    Dim tr As New ToBase64Transform
    Dim cs As CryptoStream
    Set cs = Cor.NewCryptoStream(ms, tr, WriteMode)
    Assert.IsTrue cs.CanWrite
    Assert.IsFalse cs.CanRead
End Sub

