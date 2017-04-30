VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingInfoTests
'
Option Explicit
Implements ICategorizable

Private mEncodings As New Hashtable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub FixtureSetup()
    Dim Encodings() As EncodingInfo
    Encodings = Encoding.GetEncodings
    
    Dim i As Long
    For i = LBound(Encodings) To UBound(Encodings)
        mEncodings.Add Encodings(i).CodePage, Encodings(i)
    Next
End Sub

Public Sub Equals_WithSameCodePage_ReturnsTrue()
    Dim Encoding1 As EncodingInfo
    Dim Encoding2 As EncodingInfo
    
    Set Encoding1 = mEncodings(1252&)
    Set Encoding2 = mEncodings(1252&)

    Assert.That Encoding1.Equals(Encoding2), Iz.True
    Assert.That Encoding1.GetHashCode = Encoding2.GetHashCode, Iz.True
End Sub

