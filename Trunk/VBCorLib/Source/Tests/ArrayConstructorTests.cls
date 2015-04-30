VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayConstructorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ArrayConstructorTests
'
Option Explicit

Public Sub NewChars_WithStrings_ReturnsIntegerArray()
    Dim Actual() As Integer
    
    Actual = NewChars("a", "b", "c")
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(97, 98, 99))
End Sub
