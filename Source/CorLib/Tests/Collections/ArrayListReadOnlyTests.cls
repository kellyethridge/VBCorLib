VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListReadOnlyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ReadOnly_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    ArrayList.ReadOnly Nothing
    
    AssertArgumentNullException Err, "List"
End Sub

Public Sub Add_WhenCalled_ThrowsInvalidOperationException()

End Sub
