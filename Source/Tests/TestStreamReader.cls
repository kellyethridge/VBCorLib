VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamReader
'
Option Explicit

Private ms As MemoryStream
Private sr As StreamReader
    
Const Q As Long = 10000

Public Sub Setup()
    Dim i As Long
    Dim b() As Byte
    
    
    ReDim b(Q - 1)
    Set ms = New MemoryStream
    For i = 0 To UBound(b)
        b(i) = (i Mod 10) + 48
    Next i
    ms.WriteBlock b, 0, Q
    ms.Position = 0
    Set sr = Cor.NewStreamReader(ms)
End Sub


Public Sub TestSimpleRead()
    Dim i As Long
    
    For i = 0 To 9
        Assert.AreEqual i + 48, sr.Read
    Next i
End Sub

Public Sub TestReadAll()
    Dim i As Long
    Dim j As Long
    
    i = sr.Read
    Do While i >= 0
        Assert.AreEqual (j Mod 10) + 48, i, "Wrong value at index " & j & ". "
        j = j + 1
        i = sr.Read
    Loop
    Assert.AreEqual Q, j, "10001 characters should have been read."
End Sub

Public Sub TestReadLineOf100Chars()
    ms.Position = 100
    ms.WriteByte 13
    ms.Position = 0
    
    Dim s As String
    s = sr.ReadLine
    Assert.AreEqual 100, Len(s)
End Sub

Public Sub TestRead1000TenCharLines()
    Dim i As Long
    Dim b(10) As Byte
    Dim s As String
    
    For i = 0 To 9
        b(i) = 48 + i
    Next i
    b(10) = 13
    
    ms.SetLength 0
    ms.Position = 0
    For i = 1 To 1000
        ms.WriteBlock b, 0, 11
    Next i
    
    ms.Position = 0
    
    i = 0
    s = sr.ReadLine
    Do While Not cString.IsNull(s)
        i = i + 1
        Assert.AreEqual 10, Len(s), "i = " & i
        s = sr.ReadLine
    Loop
    Assert.AreEqual 1000, i
End Sub

Public Sub TestRead10_5Kstrings()
    Dim b() As Byte
    Dim i As Long
    Dim s As String
    Dim ch() As Integer
    Dim j As Long
    
    ReDim b(0 To 5000)
    For i = 0 To 4999
        b(i) = ((i + 1) Mod 10) + 48
    Next i
    b(5000) = 13
    
    ms.SetLength 0
    ms.Position = 0
    For i = 1 To 10
        ms.WriteBlock b, 0, 5001
    Next i
    
    ms.Position = 0
    
    i = 0
    s = sr.ReadLine
    Do While Not cString.IsNull(s)
        i = i + 1
        Assert.AreEqual 5000, Len(s), "Wrong string length."
        ch = cString.ToCharArray(s)
        For j = 0 To 4999
            Assert.AreEqual b(j), ch(j)
        Next j
        s = sr.ReadLine
    Loop
    Assert.AreEqual 10, i, "Wrong number of strings read in."
End Sub

Public Sub TestReadBlock()
    Dim ch() As Integer
    Dim i As Long
    
    ReDim ch(Q - 1)
    Assert.AreEqual Q, sr.ReadBlock(ch, 0, Q)
    For i = 0 To Q - 1
        Assert.AreEqual (i Mod 10) + 48, ch(i)
    Next i
End Sub

Public Sub TestSmallPeek()
    Dim r As StreamReader
    
    Set r = NewStreamReader(App.Path & "\Dummy Test Files\3bytes.txt")
    
    Assert.IsTrue r.Peek > -1, "Peek should be greater than -1 on first peek."
    Assert.AreEqual "2" & vbCrLf, r.ReadToEnd
    Assert.AreEqual -1, r.Peek, "Peek should be -1."
    
    r.BaseStream.SeekPosition 0, FromBeginning
    r.DiscardBufferedData
    
    Assert.IsTrue r.Peek > -1, "Peek should be greater than -1 after reseting pointer."
    
    r.CloseReader

End Sub

Public Sub TestReadToEnd()
    Dim s As String
    Dim st As Stream
    Dim sr As StreamReader
    Dim sw As StreamWriter
    Dim sb As New StringBuilder
    Dim i As Long
    
    Set st = New MemoryStream
    
    For i = 0 To 10000
        sb.Append ChrW$(i And &HFFF)
    Next i
    s = sb.ToString
    
    Set sw = newstreamwriter(st)
    sw.WriteValue s
    sw.Flush
    
    st.Position = 0
    Set sr = NewStreamReader(st)
    
    Assert.AreEqual s, sr.ReadToEnd
End Sub

Public Sub TestReadLineThenReadToEnd()
    Dim st As New MemoryStream
    Dim sw As StreamWriter
    Dim sr As StreamReader
    Dim s As String
    Dim strW As New StringWriter
    
    strW.WriteLine String$(1000, ChrW$(&H2030))
    strW.WriteValue String$(9000, ChrW$(&HA0))
    
    Set sw = newstreamwriter(ms)
    sw.WriteValue strW.ToString
    sw.Flush
    
    ms.Position = 0
    Set sr = NewStreamReader(ms)
    
    Assert.AreEqual String$(1000, ChrW$(&H2030)), sr.ReadLine
    Assert.AreEqual String$(9000, ChrW$(&HA0)), sr.ReadToEnd
    
    sr.CloseReader
    sw.CloseWriter
End Sub

Public Sub TestReadEmptyLines()
    Dim ms As New MemoryStream
    Dim i As Long
    For i = 1 To 100
        ms.WriteByte 13
    Next i
    ms.Position = 0
    
    Dim sr As StreamReader
    Set sr = Cor.NewStreamReader(ms)
    
    i = 0
    Dim s As String
    Do
        s = sr.ReadLine
        If cString.IsNull(s) Then Exit Do
        Assert.AreEqual "", s
        i = i + 1
    Loop
    
    Assert.AreEqual 100, i
End Sub
