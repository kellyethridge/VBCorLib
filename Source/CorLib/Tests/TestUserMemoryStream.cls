VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUserMemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUserMemoryStream
'
Option Explicit

Private ms As MemoryStream
Private b() As Byte


Public Sub Setup()
    Dim i As Long
    
    ReDim b(11)
    For i = 0 To 11
        b(i) = i + 1
    Next i
    
    Set ms = newmemorystream(b, 2, 10)
End Sub

Public Sub Teardown()
    ms.CloseStream
End Sub

Public Sub TestLength()
    Assert.AreEqual 10@, ms.Length
End Sub

Public Sub TestReadSame()
    Dim i As Long
    
    ms.Position = 0
    For i = 0 To 9
        Assert.AreEqual b(i + 2), ms.ReadByte
    Next i
End Sub

Public Sub TestTruncateAndRestore()
    Assert.AreEqual 10@, ms.Length
    ms.SetLength 5
    Assert.AreEqual 5@, ms.Length
    ms.SetLength 10
    
    Dim i As Long
    For i = 7 To 11
        Assert.AreEqual 0, b(i)
    Next i
End Sub

Public Sub TestModifyCapacity()
    On Error GoTo errTrap
    ms.capacity = 11
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
End Sub

Public Sub TestWritePassedCapacity()
    Dim b2(10) As Byte
    
    On Error GoTo errTrap
    ms.WriteBlock b2, 0, 11
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex, Err), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is NotSupportedException, "Wrong exception type caught."
End Sub

Public Sub TestModifyRange()
    Dim i As Long
    
    ms.Position = 0
    For i = 10 To 1 Step -1
        ms.WriteByte i
    Next i
    
    ms.Position = 0
    For i = 10 To 1 Step -1
        Assert.AreEqual i, ms.ReadByte
    Next i
    For i = 2 To 11
        Assert.AreEqual 12 - i, b(i)
    Next i
End Sub
