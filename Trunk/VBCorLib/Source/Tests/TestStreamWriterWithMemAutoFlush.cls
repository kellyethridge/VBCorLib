VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSWWithMemAutoFlush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriterWithMemAutoFlush
'
Option Explicit
Implements ICategorizable

Private sw As StreamWriter
Private ms As MemoryStream

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set ms = New MemoryStream
    Set sw = newstreamwriter(ms)
    sw.AutoFlush = True
End Sub

Public Sub TestWriteLetter()
    sw.WriteValue "a"
    Assert.AreEqual 1@, ms.Length
    ms.Position = 0
    Assert.AreEqual 97, ms.ReadByte
End Sub

