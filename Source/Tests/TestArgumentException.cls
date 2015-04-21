VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArgumentException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE As String = "User Error."

Private Ex As ArgumentException
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set Ex = Cor.NewArgumentException(EXPECTED_MESSAGE, , Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.AreEqual EXPECTED_MESSAGE, Ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Set Ex = Cor.NewArgumentException(EXPECTED_MESSAGE, , Inner)
    Assert.AreEqual Inner, Ex.InnerException, "InnerException should return Inner instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As ArgumentException
    Set firstLevelEx = Cor.NewArgumentException(EXPECTED_MESSAGE, , Ex)
    Assert.AreEqual Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestUserDefinedToString()
'    Assert.IsTrue Len(Ex.ToString) > 0, "ToString return wrong value."
    Dim s As String
    s = Ex.ToString
    Assert.That s, Iz.Not.Empty
End Sub

Public Sub TestPersitence()
    Dim Copy As ArgumentException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", Ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.AreEqual EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
    'Assert.AreEqual "Argument", Copy.ParamName, "ParamName should return 'Argument'."
End Sub

'Public Sub TestParamName()
'    Assert.AreEqual "Argument", ex.ParamName, "ParamName should return 'Argument'."
'End Sub
