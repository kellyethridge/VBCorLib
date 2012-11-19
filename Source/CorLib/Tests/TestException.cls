VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE  As String = "User Error."

Public Sub Message_WhenNotSupplied_ReturnsDefault()
    Dim Ex As New Exception
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.Not.Empty
End Sub

Public Sub Message_WhenSupplied_ReturnsMessage()
    Dim Ex As Exception
    Set Ex = Cor.NewException(EXPECTED_MESSAGE)
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo(EXPECTED_MESSAGE)
End Sub

Public Sub InnerException_WhenSupplied_ReturnsInnerException()
    Dim Inner   As New Exception
    Dim Ex      As Exception
    Set Ex = Cor.NewException("", Inner)
    
    Dim Actual As Exception
    Set Actual = Ex.InnerException
    
    Assert.That Actual, Iz.SameAs(Inner)
End Sub

Public Sub GetBaseException_WithNoInnerException_ReturnsSelf()
    Dim Ex As New Exception
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Ex)
End Sub

Public Sub GetBaseException_WithMultipleInnerExceptions_ReturnsDeepestException()
    Dim Inner1 As New Exception
    Dim Inner2 As Exception
    Set Inner2 = Cor.NewException("", Inner1)
    Dim Ex As Exception
    Set Ex = Cor.NewException("", Inner2)
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Inner1)
End Sub

Public Sub TestPersitence()
    Dim Inner   As New Exception
    Dim Ex      As Exception
    Set Ex = Cor.NewException(EXPECTED_MESSAGE, Inner)
    
    Dim Bag As New PropertyBag
    Bag.WriteProperty "ex", Ex
    Dim Actual As Exception
    Set Actual = Bag.ReadProperty("ex")
    
    Assert.That Actual, Iz.Not.SameAs(Ex)
    Assert.That Actual.Message, Iz.EqualTo(EXPECTED_MESSAGE)
    Assert.That Actual.InnerException, Iz.Not.Nothing
End Sub


