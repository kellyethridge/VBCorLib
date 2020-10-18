VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileNotFoundExceptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FileNotFoundExceptionTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Message_WithDefaultConstructor_ReturnsDefaultMessage()
    Dim Ex As New FileNotFoundException
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo("Unable to find the specified file.")
End Sub

Public Sub FileName_WithDefaultConstructor_ReturnsEmpty()
    Dim Ex As New FileNotFoundException
    
    Dim Actual As String
    Actual = Ex.FileName
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub FileName_WithFileNameConstructor_ReturnsFileName()
    Dim Ex As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("abc.txt")
    
    Dim Actual As String
    Actual = Ex.FileName
    
    Assert.That Actual, Iz.EqualTo("abc.txt")
End Sub

Public Sub Message_WithFileNameConstructor_ReturnsExpected()
    Dim Ex As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("abc.txt")
    
    Dim Actual As String
    Actual = Ex.Message
    
    Assert.That Actual, Iz.EqualTo("Could not find file 'abc.txt'.")
End Sub

Public Sub Constructor_WithInnerException_StoresInnerException()
    Dim Inner   As New Exception
    Dim Ex      As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("", InnerException:=Inner)
    
    Dim Actual As Exception
    Set Actual = Ex.InnerException
    
    Assert.That Actual, Iz.SameAs(Inner)
End Sub

Public Sub GetBaseException_WithNoInnerException_ReturnsSelf()
    Dim Ex As New FileNotFoundException
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Ex)
End Sub

Public Sub GetBaseException_WithMultipleInnerExceptions_ReturnsDeepestException()
    Dim Inner1 As New Exception
    Dim Inner2 As Exception
    Set Inner2 = Cor.NewException("", InnerException:=Inner1)
    Dim Ex As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("", InnerException:=Inner2)
    
    Dim Actual As Exception
    Set Actual = Ex.GetBaseException
    
    Assert.That Actual, Iz.SameAs(Inner1)
End Sub

Public Sub Serialization_WhenCalled_SerializesException()
    Const Message As String = "File not found."
    Dim Inner   As New Exception
    Dim Ex      As FileNotFoundException
    Set Ex = Cor.NewFileNotFoundException("File", Message, Inner)
    
    Dim Actual As FileNotFoundException
    Set Actual = Persist(Ex)
    
    Assert.That Actual, Iz.Not.SameAs(Ex)
    Assert.That Actual.Message, Iz.EqualTo("File not found.")
    Assert.That Actual.FileName, Iz.EqualTo("File")
    Assert.That Actual.InnerException, Iz.Not.Nothing
End Sub
