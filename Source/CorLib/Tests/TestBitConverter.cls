VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBitConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestBitConverter
'
Option Explicit
Private b() As Byte

Public Sub TestGetBytesFromBoolean()
    b = bitconverter.GetBytes(True)
    AssertGetBytes b, 0, 1
    b = bitconverter.GetBytes(False)
    AssertGetBytes b, 0, 0
End Sub

Public Sub TestGetBytesFromInteger()
    b = bitconverter.GetBytes(-1)
    AssertGetBytes b, 1
End Sub

Public Sub TestGetBytesFromLong()
    b = bitconverter.GetBytes(-1&)
    AssertGetBytes b, 3
End Sub

Public Sub TestGetBytesFromDouble()
    Dim d As Double
    Dim dl As DLong
    ' this is to get a repeating bit pattern in a double
    dl.HiDWord = &H55555555
    dl.LoDWord = dl.HiDWord
    d = AsDouble(dl)
    
    b = bitconverter.GetBytes(d)
    AssertGetBytes b, 7, &H55
End Sub

Public Sub TestGetBytesFromCurrency()
    Dim c As Currency
    c = -0.0001@
    b = bitconverter.GetBytes(c)
    AssertGetBytes b, 7
End Sub

Public Sub TestGetBytesFromSingle()
    b = bitconverter.GetBytes(AsSingle(-1&))
    AssertGetBytes b, 3
End Sub

Public Sub TestGetBytesFromDate()
    Dim dl As DLong
    dl.HiDWord = &H1010101
    dl.LoDWord = &H1010101
    
    b = bitconverter.GetBytes(CDate(AsDouble(dl.LoDWord)))
    AssertGetBytes b, 7, 1
End Sub

Public Sub TestToBoolean()
    ReDim b(1)
    b(1) = 1
    
    Assert.IsFalse bitconverter.ToBoolean(b, 0)
    Assert.IsTrue bitconverter.ToBoolean(b, 1)
End Sub

Public Sub TestToDouble()
    b = bitconverter.GetBytes(123.45)
    Assert.AreEqual 123.45, bitconverter.ToDouble(b, 0), "ToDouble should return 123.45"
End Sub

Public Sub TestToDate()
    b = bitconverter.GetBytes(#1/1/2004 8:59:00 AM#)
    Assert.AreEqual #1/1/2004 8:59:00 AM#, bitconverter.ToDate(b, 0), "ToDate should return #1/1/2004 8:59:00 AM#."
End Sub

Public Sub TestToInteger()
    b = bitconverter.GetBytes(&H5555)
    Assert.AreEqual &H5555, bitconverter.ToInteger(b, 0), "ToInteger should return &H5555."
End Sub

Public Sub TestToLong()
    b = bitconverter.GetBytes(&H80808080)
    Assert.AreEqual &H80808080, bitconverter.ToLong(b, 0), "ToLong should return &H80808080."
End Sub

Public Sub TestToSingle()
    b = bitconverter.GetBytes(123.45!)
    Assert.AreEqual 123.45!, bitconverter.ToSingle(b, 0), "ToSingle should return 123.45."
End Sub

Public Sub TestToCurrency()
    b = bitconverter.GetBytes(1234.432@)
    Assert.AreEqual 1234.432@, bitconverter.ToCurrency(b, 0), "ToCurrency should return 1234.432"
End Sub

Public Sub TestToString()
    b = bitconverter.GetBytes(&HABCDEF01)
    Assert.AreEqual "01-EF-CD-AB", bitconverter.ToString(b), "ToString should return 01-EF-CD-AB."
End Sub

Public Sub TestToStringWithStartIndex()
    b = bitconverter.GetBytes(&HABCDEF01)
    Assert.AreEqual "CD-AB", bitconverter.ToString(b, 2), "ToString should return CD-AB."
End Sub

Public Sub TestToStringWithLength()
    b = bitconverter.GetBytes(&HABCDEF01)
    Assert.AreEqual "EF-CD", bitconverter.ToString(b, 1, 2), "ToString should return EF-CD."
End Sub

Public Sub TestToStringWithOneByte()
    ReDim b(0)
    b(0) = &H55
    Assert.AreEqual "55", bitconverter.ToString(b), "ToString should return 55."
End Sub

Public Sub TestGetBytesFromDecimal()
    Dim d As Variant
    Dim i As Long
    
    d = CDec("-1234567890.123456789")
    b = bitconverter.GetBytes(d)
    Assert.AreEqual 16, cArray.GetLength(b)
    For i = 0 To 15
        Assert.AreEqual Choose(i + 1, 21, 129, 233, 125, 244, 16, 34, 17, 0, 0, 0, 0, 0, 0, 9, 128), b(i)
    Next i
End Sub

Public Sub TestToDecimal()
    Dim d As Variant
    
    d = CDec("-1234567890.123456789")
    b = bitconverter.GetBytes(d)
    Assert.AreEqual d, bitconverter.ToDecimal(b, 0)
End Sub



Private Sub AssertGetBytes(ByRef b() As Byte, ByVal ub As Long, Optional ByVal expectedvalue As Byte = 255)
    Dim i As Long
    Assert.IsFalse cArray.IsNull(b), "A valid array should have been returned."
    Assert.AreEqual ub, UBound(b), "b UBound should be " & ub & "."
    For i = 0 To UBound(b)
        Assert.AreEqual expectedvalue, b(i), "Byte " & i & " should be " & expectedvalue & "."
    Next i
End Sub
