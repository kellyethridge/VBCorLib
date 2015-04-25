VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileNotFoundException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileNotFoundException
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub TestMessage()
    Dim Ex As New FileNotFoundException
    Assert.IsTrue Len(Ex.Message) > 0
End Sub

Public Sub TestWithFileName()
    Dim Ex As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("SomeFile")
    Assert.AreEqual "SomeFile", Ex.FileName
End Sub
