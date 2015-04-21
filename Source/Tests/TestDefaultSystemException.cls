VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDefaultSystemEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Ex As SystemException

Public Sub Setup()
    Set Ex = New SystemException
End Sub

Public Sub TestMessage()
    Assert.AreEqual "System Error.", Ex.Message, "Message should return 'An exception has occured."
End Sub

Public Sub TestNullInnerException()
    Assert.IsNothing Ex.InnerException, "InnerException should be Nothing."
End Sub

Public Sub TestGetBaseExceptionReturnsSelf()
    Assert.AreEqual Ex, Ex.GetBaseException, "GetBaseException should return itself."
End Sub

Public Sub TestToString()
    Assert.IsTrue Len(Ex.ToString), "ToString return incorrect value."
End Sub

Public Sub TestEqualsSelf()
    Assert.IsTrue Ex.Equals(Ex), "Equals should return True."
End Sub

Public Sub TestNotEqualsNonObject()
    Assert.IsFalse Ex.Equals("Not object value."), "Equals should return False."
End Sub

Public Sub TestNotEqualsNothing()
    Assert.IsFalse Ex.Equals(Nothing), "Equals should return False."
End Sub

Public Sub TestNotEqualsOther()
    Dim ex2 As New Exception
    Assert.IsFalse Ex.Equals(ex2), "Equals should return False."
End Sub

Public Sub TestGetHashCode()
    Dim ex2 As New Exception
    Assert.That ex2.GetHashCode, Iz.Not.EqualTo(Ex.GetHashCode), "GetHashCode should return different values."
End Sub



