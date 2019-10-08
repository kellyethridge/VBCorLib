Attribute VB_Name = "CorArrayCallbacks"
'
' CorArrayCallbacks
'
' Contains functions that are called during operations performed by CorArray methods.
'
Option Explicit

Public FindCallbackValue As Variant

Public Sub SetToNumber(ByRef e As Long)
    e = 5
End Sub

Public Function FindByteCallback(ByRef Value As Byte) As Boolean
    FindByteCallback = Value = CByte(FindCallbackValue)
End Function

Public Function FindIntegerCallback(ByRef Value As Integer) As Boolean
    FindIntegerCallback = Value = CInt(FindCallbackValue)
End Function

Public Function FindLongCallback(ByRef Value As Long) As Boolean
    FindLongCallback = Value = CLng(FindCallbackValue)
End Function

Public Function FindSingleCallback(ByRef Value As Single) As Boolean
    FindSingleCallback = Value = CSng(FindCallbackValue)
End Function

Public Function FindDoubleCallback(ByRef Value As Double) As Boolean
    FindDoubleCallback = Value = CDbl(FindCallbackValue)
End Function

Public Function FindStringCallback(ByRef Value As String) As Boolean
    FindStringCallback = Value = CStr(FindCallbackValue)
End Function

Public Function FindCurrencyCallback(ByRef Value As Currency) As Boolean
    FindCurrencyCallback = Value = CCur(FindCallbackValue)
End Function

Public Function FindDateCallback(ByRef Value As Date) As Boolean
    FindDateCallback = Value = CDate(FindCallbackValue)
End Function

Public Function FindVBGuidCallback(ByRef Value As CorType.VBGUID) As Boolean
    FindVBGuidCallback = Value.Data1 = CLng(FindCallbackValue)
End Function

Public Function FindInt32Callback(ByRef Value As Int32) As Boolean
    FindInt32Callback = Value.Value = CLng(FindCallbackValue)
End Function

Public Function CompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then
        CompareLongs = 1
    ElseIf x < y Then
        CompareLongs = -1
    End If
End Function

Public Function CompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then
        CompareIntegers = 1
    ElseIf x < y Then
        CompareIntegers = -1
    End If
End Function

Public Function CompareStrings(ByRef x As String, ByRef y As String) As Long
    CompareStrings = StrComp(x, y, vbBinaryCompare)
End Function

Public Function CompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then
        CompareDoubles = 1
    ElseIf x < y Then
        CompareDoubles = -1
    End If
End Function

Public Function CompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then
        CompareSingles = 1
    ElseIf x < y Then
        CompareSingles = -1
    End If
End Function

Public Function CompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then
        CompareBytes = 1
    ElseIf x < y Then
        CompareBytes = -1
    End If
End Function

Public Function CompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then
        CompareBooleans = 1
    ElseIf x < y Then
        CompareBooleans = -1
    End If
End Function

Public Function CompareDates(ByRef x As Date, ByRef y As Date) As Long
    CompareDates = DateDiff("s", y, x)
End Function

Public Function CompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then
        CompareCurrencies = 1
    ElseIf x < y Then
        CompareCurrencies = -1
    End If
End Function

Public Function CompareVariants(ByRef x As Variant, ByRef y As Variant) As Long
    CompareVariants = Comparer.Default.Compare(x, y)
End Function

Public Function CompareComparables(ByRef x As Object, ByRef y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = x
    CompareComparables = XComparable.CompareTo(y)
End Function

Public Function CompareVBGuids(ByRef x As VBGUID, ByRef y As VBGUID) As Long
    If x.Data1 < y.Data1 Then
        CompareVBGuids = -1
    ElseIf x.Data1 > y.Data1 Then
        CompareVBGuids = 1
    End If
End Function

Public Function CompareLargeUdts(ByRef x As LargeUdt, ByRef y As LargeUdt) As Long
    CompareLargeUdts = CompareLongs(x.Value, y.Value)
End Function
