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

Public Function CompareLongs(ByRef X As Long, ByRef Y As Long) As Long
    If X > Y Then
        CompareLongs = 1
    ElseIf X < Y Then
        CompareLongs = -1
    End If
End Function

Public Function CompareIntegers(ByRef X As Integer, ByRef Y As Integer) As Long
    If X > Y Then
        CompareIntegers = 1
    ElseIf X < Y Then
        CompareIntegers = -1
    End If
End Function

Public Function CompareStrings(ByRef X As String, ByRef Y As String) As Long
    CompareStrings = StrComp(X, Y, vbBinaryCompare)
End Function

Public Function CompareDoubles(ByRef X As Double, ByRef Y As Double) As Long
    If X > Y Then
        CompareDoubles = 1
    ElseIf X < Y Then
        CompareDoubles = -1
    End If
End Function

Public Function CompareSingles(ByRef X As Single, ByRef Y As Single) As Long
    If X > Y Then
        CompareSingles = 1
    ElseIf X < Y Then
        CompareSingles = -1
    End If
End Function

Public Function CompareBytes(ByRef X As Byte, ByRef Y As Byte) As Long
    If X > Y Then
        CompareBytes = 1
    ElseIf X < Y Then
        CompareBytes = -1
    End If
End Function

Public Function CompareBooleans(ByRef X As Boolean, ByRef Y As Boolean) As Long
    If X > Y Then
        CompareBooleans = 1
    ElseIf X < Y Then
        CompareBooleans = -1
    End If
End Function

Public Function CompareDates(ByRef X As Date, ByRef Y As Date) As Long
    CompareDates = DateDiff("s", Y, X)
End Function

Public Function CompareCurrencies(ByRef X As Currency, ByRef Y As Currency) As Long
    If X > Y Then
        CompareCurrencies = 1
    ElseIf X < Y Then
        CompareCurrencies = -1
    End If
End Function

Public Function CompareVariants(ByRef X As Variant, ByRef Y As Variant) As Long
    CompareVariants = Comparer.Default.Compare(X, Y)
End Function

Public Function CompareComparables(ByRef X As Object, ByRef Y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = X
    CompareComparables = XComparable.CompareTo(Y)
End Function

Public Function CompareVBGuids(ByRef X As VBGUID, ByRef Y As VBGUID) As Long
    If X.Data1 < Y.Data1 Then
        CompareVBGuids = -1
    ElseIf X.Data1 > Y.Data1 Then
        CompareVBGuids = 1
    End If
End Function
