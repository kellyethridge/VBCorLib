Attribute VB_Name = "BigIntegerAssertions"
Option Explicit

Public Sub AssertNumber(ByVal Number As BigInteger, ByRef ExpectedBytes() As Byte, ByVal ExpectedSign As Long)
    Assert.That Number.ToByteArray, Iz.EqualTo(ExpectedBytes)
    Assert.That Number.Sign, Iz.EqualTo(ExpectedSign)
End Sub
