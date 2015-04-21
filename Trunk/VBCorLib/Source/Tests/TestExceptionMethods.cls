VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestExceptionMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestThrowCatch()
    Dim original As ArgumentException
    Set original = New ArgumentException
    On Error GoTo errTrap
    Throw original
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim Ex As SystemException
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.AreEqual original, Ex, "Should have caught original instance."
End Sub

Public Sub TestNothingToCatch()
    ClearException
    
    Dim Ex As Exception
    Assert.IsFalse Catch(Ex), "No exception should have been caught."
End Sub

Public Sub TestCatchErrRaise()
    Dim Ex As Exception
    
    On Error GoTo errTrap
    Err.Raise 383, "TestCatchErrRaise", "Test Catch."
    On Error GoTo 0
    Assert.Fail "An error should have been thrown."
errTrap:
    Assert.IsTrue Catch(Ex, Err), "An error should have been caught."
    Assert.AreEqual 383, Ex.HResult, "HResult should be 383."
    Assert.AreEqual "TestCatchErrRaise", Ex.Source, "Source should be 'TestCatchErrRaise'."
    Assert.AreEqual "Test Catch.", Ex.Message, "Message should be 'Test Catch.'"
End Sub

Public Sub TestReThrow()
    On Error GoTo errTrap
    ReThrow
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.AreEqual "Rethrown", Ex.Message
End Sub

Public Sub TestNothingToReThrow()
    ClearException
    
    On Error GoTo errTrap
    Throw
    Exit Sub
    
errTrap:
    Assert.Fail "No exception should be thrown."
End Sub

Public Sub TestReThrowErrObject()
    On Error GoTo errTrap
    ReThrowErrObject
    Exit Sub
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.AreEqual 123, Ex.HResult
    Assert.AreEqual "Local", Ex.Source
    Assert.AreEqual "Thrown", Ex.Message
    Assert.AreEqual "www.help.com", Ex.HelpLink
End Sub



Private Sub ReThrowErrObject()
    On Error GoTo errTrap
    Err.Raise 123, "Local", "Thrown", "www.help.com"
    Exit Sub
    
errTrap:
    Throw Err
End Sub

Private Sub ReThrow()
    On Error GoTo errTrap
    Throw Cor.NewException("Rethrown")
    Exit Sub
    
errTrap:
    Throw
End Sub
