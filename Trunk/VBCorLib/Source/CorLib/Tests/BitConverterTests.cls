VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitConverterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestBitConverter
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithValue_ReturnsExpected"
            Test.Use True, NewBytes(1)
            Test.Use False, NewBytes(0)
            Test.Use 1, NewBytes(1, 0)
            Test.Use -1, NewBytes(255, 255)
            Test.Use &H8000, NewBytes(&H0, &H80)
            Test.Use 1&, NewBytes(1, 0, 0, 0)
            Test.Use -1&, NewBytes(255, 255, 255, 255)
            Test.Use &H80000000, NewBytes(&H0, &H0, &H0, &H80)
            Test.Use 1#, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &HF0, &H3F)
            Test.Use -1#, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &HF0, &HBF)
            Test.Use 1.2345E+104, NewBytes(&H38, &H16, &HCA, &HF4, &H16, &H8F, &H8B, &H55)
            Test.Use 1!, NewBytes(&H0, &H0, &H80, &H3F)
            Test.Use -1!, NewBytes(&H0, &H0, &H80, &HBF)
            Test.Use 1.2345E+10!, NewBytes(&H70, &HF4, &H37, &H50)
            Test.Use 0.0001@, NewBytes(1, 0, 0, 0, 0, 0, 0, 0)
            Test.Use -922337203685477.5807@, NewBytes(&H1, &H0, &H0, &H0, &H0, &H0, &H0, &H80)
            Test.Use 922337203685477.5807@, NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H7F)
            Test.Use #1/2/2001 8:30:45 AM#, NewBytes(&H9A, &H99, &H99, &H59, &HAB, &H3, &HE2, &H40)
            Test.Use CDec("12345.678911"), NewBytes(&H3F, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H0)
            Test.Use CDec("-12345.678911"), NewBytes(&H3F, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H80)
            
    End Select
End Sub

Public Sub GetBytes_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Variant)
    Dim Actual() As Byte
    
    Actual = BitConverter.GetBytes(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'Public Sub TestToBoolean()
'    ReDim b(1)
'    b(1) = 1
'
'    Assert.IsFalse bitconverter.ToBoolean(b, 0)
'    Assert.IsTrue bitconverter.ToBoolean(b, 1)
'End Sub
'
'Public Sub TestToDouble()
'    b = bitconverter.GetBytes(123.45)
'    Assert.AreEqual 123.45, bitconverter.ToDouble(b, 0), "ToDouble should return 123.45"
'End Sub
'
'Public Sub TestToDate()
'    b = bitconverter.GetBytes(#1/1/2004 8:59:00 AM#)
'    Assert.AreEqual #1/1/2004 8:59:00 AM#, bitconverter.ToDate(b, 0), "ToDate should return #1/1/2004 8:59:00 AM#."
'End Sub
'
'Public Sub TestToInteger()
'    b = bitconverter.GetBytes(&H5555)
'    Assert.AreEqual &H5555, bitconverter.ToInteger(b, 0), "ToInteger should return &H5555."
'End Sub
'
'Public Sub TestToLong()
'    b = bitconverter.GetBytes(&H80808080)
'    Assert.AreEqual &H80808080, bitconverter.ToLong(b, 0), "ToLong should return &H80808080."
'End Sub
'
'Public Sub TestToSingle()
'    b = bitconverter.GetBytes(123.45!)
'    Assert.AreEqual 123.45!, bitconverter.ToSingle(b, 0), "ToSingle should return 123.45."
'End Sub
'
'Public Sub TestToCurrency()
'    b = bitconverter.GetBytes(1234.432@)
'    Assert.AreEqual 1234.432@, bitconverter.ToCurrency(b, 0), "ToCurrency should return 1234.432"
'End Sub
'
'Public Sub TestToString()
'    b = bitconverter.GetBytes(&HABCDEF01)
'    Assert.AreEqual "01-EF-CD-AB", bitconverter.ToString(b), "ToString should return 01-EF-CD-AB."
'End Sub
'
'Public Sub TestToStringWithStartIndex()
'    b = bitconverter.GetBytes(&HABCDEF01)
'    Assert.AreEqual "CD-AB", bitconverter.ToString(b, 2), "ToString should return CD-AB."
'End Sub
'
'Public Sub TestToStringWithLength()
'    b = bitconverter.GetBytes(&HABCDEF01)
'    Assert.AreEqual "EF-CD", bitconverter.ToString(b, 1, 2), "ToString should return EF-CD."
'End Sub
'
'Public Sub TestToStringWithOneByte()
'    ReDim b(0)
'    b(0) = &H55
'    Assert.AreEqual "55", bitconverter.ToString(b), "ToString should return 55."
'End Sub
'
'Public Sub TestToDecimal()
'    Dim d As Variant
'
'    d = CDec("-1234567890.123456789")
'    b = bitconverter.GetBytes(d)
'    Assert.AreEqual d, bitconverter.ToDecimal(b, 0)
'End Sub
