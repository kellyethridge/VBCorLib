VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDetermineEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamReaderDetermineEncoding
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub TestUTF8Encoding()
    TestEncoding Cor.NewUTF8Encoding(True, False)
End Sub

Public Sub TestUnicodeEnocding()
    TestEncoding Cor.NewUnicodeEncoding(False, True)
End Sub

Public Sub TestUnicodeEncodingBigEndian()
    TestEncoding Cor.NewUnicodeEncoding(True, True)
End Sub



Private Sub TestEncoding(ByVal En As Encoding)
    Dim ms As MemoryStream
    Dim sr As StreamReader
    Dim b() As Byte
    Dim currEn As Encoding
    
    Set ms = New MemoryStream
    ms.WriteBlock En.GetPreamble, 0, CorArray.GetLength(En.GetPreamble)
    b = En.GetBytes(ChrW$(55))
    ms.WriteBlock b, 0, CorArray.GetLength(b)
    ms.Position = 0
    
    Set sr = Cor.NewStreamReader(ms, , True)
    Set currEn = sr.CurrentEncoding
    
    Assert.AreEqual 55, sr.Read
    Assert.IsFalse sr.CurrentEncoding Is currEn
    Assert.AreEqual TypeName(En), TypeName(sr.CurrentEncoding)
End Sub
