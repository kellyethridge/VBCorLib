VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSystemException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable

Private Const EXPECTED_MESSAGE As String = "User Error."
Private Const MESSAGE_PREFIX As String = "VBCorLib.SystemException: "

Private Ex As SystemException
Private Inner As Exception

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set Inner = New Exception
    Set Ex = Cor.NewSystemException(EXPECTED_MESSAGE, Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.AreEqual EXPECTED_MESSAGE, Ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Assert.AreEqual Inner, Ex.InnerException, "InnerException should return UserException instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As SystemException
    Set firstLevelEx = Cor.NewSystemException(EXPECTED_MESSAGE, Ex)
    Assert.AreEqual Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.IsTrue Len(Ex.ToString) > 0, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim Copy As SystemException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", Ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.AreEqual EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
End Sub

