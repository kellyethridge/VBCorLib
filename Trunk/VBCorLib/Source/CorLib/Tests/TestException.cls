VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE  As String = "User Error."

Private Ex As Exception
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set Ex = cor.NewException(EXPECTED_MESSAGE, Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.AreEqual EXPECTED_MESSAGE, Ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Assert.AreEqual Inner, Ex.InnerException, "InnerException should return Inner instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As Exception
    Set firstLevelEx = cor.NewException(EXPECTED_MESSAGE, Ex)
    Assert.AreEqual Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestPersitence()
    Dim Copy As Exception
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", Ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.That Ex, Iz.Not.EqualTo(Copy), "copy should be a different object instance."
    Assert.AreEqual EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."

End Sub


