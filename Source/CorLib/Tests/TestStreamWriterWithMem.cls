VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamWriterWithMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriterWithMem
'
Option Explicit

Private sw As StreamWriter
Private ms As MemoryStream

Public Sub Setup()
    Set ms = New MemoryStream
    Set sw = cor.newstreamwriter(ms)
End Sub

Public Sub TestWriteLetter()
    sw.WriteValue "a"
    Assert.AreEqual 0@, ms.Length
    sw.Flush
    Assert.AreEqual 1@, ms.Length
    ms.Position = 0
    Assert.AreEqual 97, ms.ReadByte
End Sub

Public Sub TestWriteNumbersWithLineBreaks()
    Dim i As Long
    
    For i = 1 To 100
        sw.WriteLine i
    Next i
    sw.WriteLine
    sw.Flush
    
    Assert.AreEqual 394@, ms.Length, "Stream length should be 392."
    
    ms.Position = 0
    Dim sr As StreamReader
    Set sr = NewStreamReader(ms)
    
    For i = 1 To 100
        Assert.AreEqual i, CLng(sr.ReadLine)
    Next i
    
    Assert.That 0, Iz.Not.EqualTo(StrPtr(sr.ReadLine))
    Assert.AreEqual 0, StrPtr(sr.ReadLine)
End Sub
