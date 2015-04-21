Attribute VB_Name = "modSorterCallbacks"
Option Explicit

Public Function CompareStringsAscending(ByRef X As String, ByRef Y As String) As Long
    CompareStringsAscending = StrComp(X, Y)
End Function

Public Function CompareVBGuids(ByRef X As VBGUID, ByRef Y As VBGUID) As Long
    If X.Data1 < Y.Data1 Then
        CompareVBGuids = -1
    ElseIf X.Data1 > Y.Data1 Then
        CompareVBGuids = 1
    End If
End Function
