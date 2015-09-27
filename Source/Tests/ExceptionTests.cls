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
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Message_WithDefaultConstructor_ReturnsDefaultMessage()
    Dim Ex As New Exception
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo("Exception of type 'CorLib.Exception' was thrown.")
End Sub

Public Sub Constructor_WithMessage_StoresMessage()
    Const Expected As String = "User Error."
    Dim Ex As Exception
    Set Ex = Cor.NewException(Expected)
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub InnerException_WhenSupplied_ReturnsInnerException()
    Dim Inner   As New Exception
    Dim Ex      As Exception
    Set Ex = Cor.NewException("", InnerException:=Inner)
    
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
    Set Inner2 = Cor.NewException("", InnerException:=Inner1)
    Dim Ex As Exception
    Set Ex = Cor.NewException("", InnerException:=Inner2)
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Inner1)
End Sub

Public Sub Serialization_WhenCalled_ClonesException()
    Dim Inner   As New Exception
    Dim Ex      As Exception
    Set Ex = Cor.NewException("Msg", 91, Inner)
    Ex.Source = "abc"
    Ex.HelpLink = "Link"
    Ex.Data("x") = 1
    
    Dim Actual As Exception
    Set Actual = Persist(Ex)
    
    Assert.That Actual, Iz.Not.SameAs(Ex)
    Assert.That Actual.Message, Iz.EqualTo("Msg")
    Assert.That Actual.InnerException, Iz.Not.Nothing
    Assert.That Actual.Source, Iz.EqualTo("abc")
    Assert.That Actual.HelpLink, Iz.EqualTo("Link")
    Assert.That Actual.Data("x"), Iz.EqualTo(1)
    Assert.That Actual.ErrorNumber, Iz.EqualTo(91)
End Sub


