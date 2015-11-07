VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvalidCastExceptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' InvalidCastExceptionTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Message_WithDefaultConstructor_ReturnsDefaultMessage()
    Dim Ex As New InvalidCastException
    
    Dim Actual As String
    Actual = Ex.Message
    
    Const Expected As String = "Specified cast is not valid."
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Constructor_WithMessage_StoresMessage()
    Const Expected As String = "User Error."
    Dim Ex As InvalidCastException
    Set Ex = CreateException(Expected)
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub InnerException_WhenSupplied_ReturnsInnerException()
    Dim Inner   As New Exception
    Dim Ex      As InvalidCastException
    Set Ex = CreateException("", Inner)
    
    Dim Actual As Exception
    Set Actual = Ex.InnerException
    
    Assert.That Actual, Iz.SameAs(Inner)
End Sub

Public Sub GetBaseException_WithNoInnerException_ReturnsSelf()
    Dim Ex As New InvalidCastException
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Ex)
End Sub

Public Sub GetBaseException_WithMultipleInnerExceptions_ReturnsDeepestException()
    Dim Inner1 As New Exception
    Dim Inner2 As Exception
    Set Inner2 = CreateException("", Inner1)
    Dim Ex As InvalidCastException
    Set Ex = CreateException("", Inner2)
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Inner1)
End Sub

Public Sub Serialization_WhenCalled_SerializesException()
    Const Message As String = "User Error."
    Dim Inner   As New Exception
    Dim Ex      As InvalidCastException
    Set Ex = CreateException(Message, Inner)
    
    Dim Actual As Exception
    Set Actual = Persist(Ex)
    
    Assert.That Actual, Iz.Not.SameAs(Ex)
    Assert.That Actual.Message, Iz.EqualTo(Message)
    Assert.That Actual.InnerException, Iz.Not.Nothing
End Sub

Private Function CreateException(Optional ByRef Message As String, Optional ByVal Inner As Exception) As Exception
    Set CreateException = Cor.NewInvalidCastException(Message, Inner)
End Function



