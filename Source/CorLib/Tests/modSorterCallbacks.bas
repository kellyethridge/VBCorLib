Attribute VB_Name = "modSorterCallbacks"
Option Explicit

Public Function CompareStrings(ByRef x As String, ByRef y As String) As Long
    CompareStrings = StrComp(x, y)
End Function

Public Function CompareVBGuids(ByRef x As VBGUID, ByRef y As VBGUID) As Long
    If x.Data1 < y.Data1 Then
        CompareVBGuids = -1
    ElseIf x.Data1 > y.Data1 Then
        CompareVBGuids = 1
    End If
End Function
