VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestInvalidCastException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable

Private Const EXPECTED_MESSAGE As String = "User Error."

Private Ex As InvalidCastException
Private Inner As Exception

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set Inner = New Exception
    Set Ex = Cor.NewInvalidCastException(EXPECTED_MESSAGE, Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.AreEqual EXPECTED_MESSAGE, Ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Set Ex = Cor.NewInvalidCastException(EXPECTED_MESSAGE, Inner)
    Assert.AreEqual Inner, Ex.InnerException, "InnerException should return Inner instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As InvalidCastException
    Set firstLevelEx = Cor.NewInvalidCastException(EXPECTED_MESSAGE, Ex)
    Assert.AreEqual Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.IsTrue InStr(Ex.ToString, EXPECTED_MESSAGE) > 0, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim Copy As InvalidCastException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", Ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.AreEqual EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
End Sub


