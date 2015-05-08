Attribute VB_Name = "cArrayCallbacks"
'
' CorArrayCallbacks
'
' Contains functions that are called during operations performed by CorArray methods.
'
Option Explicit

Public FindCallbackValue As Variant

Public Function MakeInt32(ByVal Value As Long) As Int32
    Set MakeInt32 = New Int32
    MakeInt32.Value = Value
End Function

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

Public Function FindVBGuidCallback(ByRef Value As VBCorType.VBGUID) As Boolean
    FindVBGuidCallback = Value.Data1 = CLng(FindCallbackValue)
End Function

Public Function FindInt32Callback(ByRef Value As Int32) As Boolean
    FindInt32Callback = Value.Value = CLng(FindCallbackValue)
End Function

