VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoStreamMultiBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoStreamMultiBlock
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Write100Bytes()
    Dim csp As New DESCryptoServiceProvider
    Dim cs As CryptoStream
    Dim ms As New MemoryStream
    Set cs = NewCryptoStream(ms, csp.CreateEncryptor, WriteMode)
    
    Dim b() As Byte
    ReDim b(99)
    
    Dim r As New Random
    r.NextBytes b
    
    cs.WriteBlock b, 0, 100
    cs.FlushFinalBlock
    
    Assert.AreEqual 104, ms.Length
    
    Dim b2() As Byte
    ReDim b2(99)
    
    Dim ms2 As New MemoryStream
    ms.WriteTo ms2
    ms2.Position = 0
    Set cs = NewCryptoStream(ms2, csp.CreateDecryptor, ReadMode)
    
    Assert.AreEqual 100, cs.ReadBlock(b2, 0, 100), "Wrong readblock length."
    Assert.AreEqual b, b2
End Sub

Public Sub Write8Bytes()
    Dim csp As New DESCryptoServiceProvider
    Dim cs As CryptoStream
    Dim ms As New MemoryStream
    Set cs = NewCryptoStream(ms, csp.CreateEncryptor, WriteMode)
    
    Dim b() As Byte
    ReDim b(7)
    
    Dim r As New Random
    r.NextBytes b
    
    cs.WriteBlock b, 0, 8
    cs.FlushFinalBlock
    
    Assert.AreEqual 16, ms.Length
    
    Dim b2() As Byte
    ReDim b2(7)
    
    Dim ms2 As New MemoryStream
    ms.WriteTo ms2
    ms2.Position = 0
    Set cs = NewCryptoStream(ms2, csp.CreateDecryptor, ReadMode)
    
    Assert.AreEqual 8, cs.ReadBlock(b2, 0, 8), "Wrong readblock length."
    Assert.AreEqual b, b2
End Sub

Public Sub LongTestWriteDifferentBlockLengths()
    Dim csp As New DESCryptoServiceProvider
    csp.Padding = ANSIX923
    
    Dim cs As CryptoStream
    Dim ms As MemoryStream
    Dim ms2 As MemoryStream
    
    Dim b() As Byte
    ReDim b(1000)
    
    Dim r As New Random
    r.NextBytes b
    
    ' 1 to 100 byte block lengths will be written.
    Dim i As Long
    For i = 1 To 100
        Set ms = New MemoryStream
        Set cs = NewCryptoStream(ms, csp.CreateEncryptor, WriteMode)
        Dim c As Long
        c = 0
        Do While (CorArray.GetLength(b) - c) > i
            cs.WriteBlock b, c, i
            c = c + i
        Loop
        
        If c < CorArray.GetLength(b) Then
            cs.WriteBlock b, c, CorArray.GetLength(b) - c
        End If
        cs.FlushFinalBlock
        
        Assert.AreEqual 1008, ms.Length
        
        Set ms2 = New MemoryStream
        ms.WriteTo ms2
        ms2.Position = 0
        Set cs = NewCryptoStream(ms2, csp.CreateDecryptor, ReadMode)
        Dim b2() As Byte
        ReDim b2(0 To CorArray.GetLength(b) - 1)
        Assert.AreEqual CorArray.GetLength(b), cs.ReadBlock(b2, 0, CorArray.GetLength(b))
        Assert.AreEqual b, b2
    Next i
End Sub

Public Sub LongTestReadDifferentBlockLengths()
    Dim csp As New DESCryptoServiceProvider
    csp.Padding = ISO10126
    
    Dim cs As CryptoStream
    Dim ms As New MemoryStream
    Dim ms2 As New MemoryStream
    
    Dim b() As Byte
    ReDim b(1000)
    
    Dim r As New Random
    r.NextBytes b
    
    Set cs = NewCryptoStream(ms, csp.CreateEncryptor, WriteMode)
    cs.WriteBlock b, 0, CorArray.GetLength(b)
    cs.FlushFinalBlock
    ms.WriteTo ms2
    
    Dim i As Long
    Dim b2() As Byte
    ReDim b2(1000)
    
    For i = 1 To 100
        Set ms = New MemoryStream
        ms2.WriteTo ms
        ms.Position = 0
        Set cs = NewCryptoStream(ms, csp.CreateDecryptor, ReadMode)
        Dim c As Long
        c = 0
        Do While (CorArray.GetLength(b2) - c) > i
            Assert.AreEqual i, cs.ReadBlock(b2, c, i), "Read failed."
            c = c + i
        Loop
        
        If c < CorArray.GetLength(b2) Then
            Assert.AreEqual CorArray.GetLength(b2) - c, cs.ReadBlock(b2, c, CorArray.GetLength(b2) - c), "Final block read failed."
        End If
        
        Assert.AreEqual b, b2, "Array is wrong."
    Next i
End Sub

Public Sub BigRead()
    Dim csp As New DESCryptoServiceProvider
    csp.Padding = ISO10126
    
    Dim cs As CryptoStream
    Dim ms As New MemoryStream
    
    Dim b() As Byte
    ReDim b(1000)
    
    Dim r As New Random
    r.NextBytes b
    
    Set cs = NewCryptoStream(ms, csp.CreateEncryptor, WriteMode)
    cs.WriteBlock b, 0, CorArray.GetLength(b)
    cs.FlushFinalBlock
    
    Dim ms2 As New MemoryStream
    ms.WriteTo ms2
    ms2.Position = 0
    
    Set cs = NewCryptoStream(ms2, csp.CreateDecryptor, ReadMode)
    
    Dim b2() As Byte
    ReDim b2(2000)
    Assert.AreEqual 1001, cs.ReadBlock(b2, 0, 1100)
    Assert.AreEqual 0, cs.ReadBlock(b2, 0, 1)
    ReDim Preserve b2(1000)
    Assert.AreEqual b, b2
End Sub

