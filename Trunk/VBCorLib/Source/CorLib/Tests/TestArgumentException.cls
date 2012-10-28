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

Private ex As ArgumentException
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set ex = Cor.NewArgumentException(EXPECTED_MESSAGE, , Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.AreEqual EXPECTED_MESSAGE, ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Set ex = Cor.NewArgumentException(EXPECTED_MESSAGE, , Inner)
    Assert.AreEqual Inner, ex.InnerException, "InnerException should return Inner instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As ArgumentException
    Set firstLevelEx = Cor.NewArgumentException(EXPECTED_MESSAGE, , ex)
    Assert.AreEqual Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.IsTrue Len(ex.ToString) > 0, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim Copy As ArgumentException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.AreEqual EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
    'Assert.AreEqual "Argument", Copy.ParamName, "ParamName should return 'Argument'."
End Sub

'Public Sub TestParamName()
'    Assert.AreEqual "Argument", ex.ParamName, "ParamName should return 'Argument'."
'End Sub
