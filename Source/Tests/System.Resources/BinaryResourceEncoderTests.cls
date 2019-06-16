VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryResourceEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BinaryResourceEncoderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Encode_WithNonByteArray_ReturnsFalse()
    Dim Encoder As New BinaryResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode("abc", "name")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Encode_WithNullByteArray_ReturnsFalse()
    Dim Encoder As New BinaryResourceEncoder
    Dim Bytes() As Byte
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(Bytes, "name")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Encode_WithByteArray_ReturnsTrue()
    Dim Encoder As New BinaryResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(NewBytes(1), "Bytes")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WithDefaultResourceType_ReturnsResourceKey()
    Dim Encoder As New BinaryResourceEncoder
    Encoder.Encode NewBytes(1), "Bytes"
    
    Dim ActualKey As ResourceKey
    Dim ActualValue() As Byte
    Assert.That Encoder.GetEncodedResource(ActualKey, ActualValue), Iz.True
    
    Assert.That ActualKey.ResourceName, Iz.EqualTo("Bytes")
    Assert.That ActualKey.ResourceType, Iz.EqualTo("CUSTOM")
End Sub

Public Sub GetEncodedResource_WithSpecifiedResourceType_ReturnsResourceKey()
    Dim Encoder As New BinaryResourceEncoder
    Encoder.Encode NewBytes(1), "Bytes", 1234
    
    Dim ActualKey As ResourceKey
    Dim ActualValue() As Byte
    Assert.That Encoder.GetEncodedResource(ActualKey, ActualValue), Iz.True
    
    Assert.That ActualKey.ResourceName, Iz.EqualTo("Bytes")
    Assert.That ActualKey.ResourceType, Iz.EqualTo(1234)
End Sub

Public Sub GetEncodedResource_WhenCalled_ReturnsValue()
    Dim Encoder As New BinaryResourceEncoder
    Encoder.Encode NewBytes(1), "Bytes", 1234
    
    Dim ActualKey As ResourceKey
    Dim ActualValue() As Byte
    Assert.That Encoder.GetEncodedResource(ActualKey, ActualValue), Iz.True

    Assert.That ActualValue, Iz.EqualTo(NewBytes(1))
End Sub

Public Sub GetEncodedResource_WhenAllResourcedReturns_ReturnsFalse()
    Dim Encoder As New BinaryResourceEncoder
    Dim Key As ResourceKey
    Dim Value() As Byte
    Encoder.Encode NewBytes(1), "Bytes", 1234
    Encoder.GetEncodedResource Key, Value

    Dim Actual As Boolean
    Actual = Encoder.GetEncodedResource(Key, Value)
    
    Assert.That Actual, Iz.False
End Sub
