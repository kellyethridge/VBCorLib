VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCharEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCharEnumerator
'
Option Explicit

Public Sub TestManualIteration()
    Dim s As String
    Dim en As CharEnumerator
    Dim i As Long
    
    s = "hello"
    Set en = Cor.NewCharEnumerator(s)
    
    Do While en.MoveNext
        i = i + 1
        Assert.AreEqual Choose(i, 104, 101, 108, 108, 111), en.Current
    Loop
End Sub

Public Sub TestForEachIteration()
    Dim s As String
    Dim i As Long
    Dim ch As Variant
    
    s = "hello"
    
    For Each ch In Cor.NewCharEnumerator(s)
        i = i + 1
        Assert.AreEqual Choose(i, 104, 101, 108, 108, 111), ch
    Next ch
End Sub
