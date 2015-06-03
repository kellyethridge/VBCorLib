VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentExceptionTests"
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
    Dim Ex As New ArgumentException
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo("Value does not fall within the expected range.")
End Sub

Public Sub ParamName_WithDefaultConstructor_ReturnsEmpty()
    Dim Ex As New ArgumentException
    
    Dim Actual As String
    Actual = Ex.ParamName
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub Constructor_WithMessage_StoresMessage()
    Const Expected As String = "User Error."
    Dim Ex As ArgumentException
    Set Ex = Cor.NewArgumentException(Expected)
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Constructor_WithParamName_AppendsToMessage()
    Dim Ex As ArgumentException
    Set Ex = Cor.NewArgumentException("Message", "Param")
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo("Message" & vbCrLf & "Parameter name: Param")
End Sub

Public Sub Constructor_WithInnerException_StoresInnerException()
    Dim Inner   As New Exception
    Dim Ex      As ArgumentException
    Set Ex = Cor.NewArgumentException("", , Inner)
    
    Dim Actual As Exception
    Set Actual = Ex.InnerException
    
    Assert.That Actual, Iz.SameAs(Inner)
End Sub

Public Sub GetBaseException_WithNoInnerException_ReturnsSelf()
    Dim Ex As New ArgumentException
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Ex)
End Sub

Public Sub GetBaseException_WithMultipleInnerExceptions_ReturnsDeepestException()
    Dim Inner1 As New Exception
    Dim Inner2 As Exception
    Set Inner2 = Cor.NewException("", InnerException:=Inner1)
    Dim Ex As ArgumentException
    Set Ex = Cor.NewArgumentException("", InnerException:=Inner2)
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Inner1)
End Sub

Public Sub TestPersitence()
    Const Message As String = "User Error."
    Dim Inner   As New Exception
    Dim Ex      As ArgumentException
    Set Ex = Cor.NewArgumentException(Message, "Param", Inner)
    
    Dim Bag As New PropertyBag
    Bag.WriteProperty "ex", Ex
    Dim Actual As ArgumentException
    Set Actual = Bag.ReadProperty("ex")
    
    Assert.That Actual, Iz.Not.SameAs(Ex)
    Assert.That Actual.Message, Iz.EqualTo(Message & vbCrLf & "Parameter name: Param")
    Assert.That Actual.ParamName, Iz.EqualTo("Param")
    Assert.That Actual.InnerException, Iz.Not.Nothing
End Sub

