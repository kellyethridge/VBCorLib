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

Public Sub TestMessage()
    Dim ex As New FileNotFoundException
    Assert.IsTrue Len(ex.Message) > 0
End Sub

Public Sub TestWithFileName()
    Dim ex As FileNotFoundException
    Set ex = Cor.NewFileNotFoundException("SomeFile")
    Assert.AreEqual "SomeFile", ex.FileName
End Sub
