VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringResourceEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringResourceEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Encode_WithNonStringValue_ReturnsFalse"
            Test.Use 1
            Test.Use New StdPicture
            Test.Use Empty
            Test.Use Null
            Test.Use Now
        
        Case "Encode_WithNonNumericResourceName_ReturnsFalse"
            Test.Use "1"
            Test.Use New StdPicture
            Test.Use Empty
            Test.Use Null
            Test.Use Now
            
    End Select
End Sub

Public Sub Encode_WithNonStringValue_ReturnsFalse(ByRef NonStringValue As Variant)
    Dim Encoder As New StringResourceEncoder
    
    Assert.That Encoder.Encode(NonStringValue, 1), Iz.False
End Sub

Public Sub Encode_WithNonNumericResourceName_ReturnsFalse(ByRef NonNumericValue As Variant)
    Dim Encoder As New StringResourceEncoder
    
    Assert.That Encoder.Encode("abc", NonNumericValue), Iz.False
End Sub

Public Sub GetEncodedResource_WithStringsEncoded_ReturnsEncodedData()
    Dim Encoder As New StringResourceEncoder
    
    Encoder.Encode "a", 1
    Encoder.Encode "b", 2
    
    Dim Key As ResourceKey
    Dim Data() As Byte
    Encoder.GetEncodedResource Key, Data
    
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.StringResource)
    Assert.That Key.ResourceName, Iz.EqualTo(1)
    Dim Expected() As Byte
    Expected = NewBytes(0, 0, 1, 0, &H61, 0, 1, 0, &H62, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    Assert.That Data, Iz.EqualTo(Expected)
End Sub
