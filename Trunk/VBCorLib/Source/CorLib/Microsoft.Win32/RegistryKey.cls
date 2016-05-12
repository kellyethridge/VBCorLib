VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: RegistryKey
'

''
' Represents a key level node inside the Windows Registry.
'
' @remarks This class allows for easy access and manipulation of keys and key values
' within the Windows Registry. By using a key naming convention similar to folders the
' keys and values can be traversed and modified.
' <p>There a set of Root nodes that separate the Registry to common areas for specific
' needs. The Root nodes are the starting place for any keys in the registry. A key name
' could look like 'HKEY_LOCAL_MACHINE\Software'.
' <pre>
' Iterate the values found in a registry key.
'
' Private Sub Main()
'     Dim rk As RegistryKey
'     Dim names() As String
'     Dim i As Long
'
'     '' Open a registry subkey for enumeration.
'     Set rk = Registry.LocalMachine.OpenSubKey("software\microsoft\windows\currentversion\run")
'
'     '' Retrieve all names for the values in the key.
'     names = rk.GetValueNames
'
'     '' enumerate the names and get the value for each,
'     '' displaying the pair as [name] = [value].
'     For i = 0 To UBound(names)
'         Console.WriteLine "{0} = {1}", names(i), rk.GetValue(names(i))
'     Next i
'
'     '' Close the registry key.
'     rk.CloseKey
'
'     '' Wait for user to hit return.
'     Console.ReadLine
' End Sub
'
' ''This code produces the following output.
' ''The list will vary from machine to machine.
'
' ''CplBCL50 = C:\Program Files\EzButton\CplBCL50.EXE
' ''ccApp = "C:\Program Files\Common Files\Symantec Shared\ccApp.exe"
' ''Symantec NetDriver Monitor = C:\PROGRA~1\SYMNET~1\SNDMon.exe
' </pre>
'
' @see Registry
' @see RegistryKeyStatic
'
Option Explicit
Implements IObject

Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const DEFAULT_VALUE_NAME        As String = ""
Private Const MaxKeyLength              As Long = 255
Private Const MaxValueNameLength        As Long = 16383


''
' The value types that can be set and retrieved from the Registry.
'
' @param UnknownKind Determines the value type based on the variable type or registry setting.
' @param StringKind Attempts to set or retrieve the value as a string.
' @param DWordKind Attempts to set or retrieve the value as a DWord (Long).
' @param BinaryKind Attempts to set or retrieve the value as a Byte array.
' @param MultiStringKind Attempts to set or retrieve the value as a String array.
' @param ExpandStringKind Attempts to set or retrieve the value as an expandable string.
' @param QWordKind Attempts to set or retrieve the value as a 64-bit value.
'
Public Enum RegistryValueKind
    UnknownKind = 0
    StringKind = 1
    DWordKind = 4
    BinaryKind = 3
    MultiStringKind = 7
    ExpandStringKind = 2
    QWordKind = 11
End Enum

Public Enum RegistryValueOptions
    None = 0
    DoNotExpandEnvironmentVariables = 1
End Enum

Private mHKey       As Long
Private mName       As String
Private mWritable   As Boolean


''
' Returns the name of the key.
'
' @return The name of the key.
'
Public Property Get Name() As String
    EnsureNotDisposed
    Name = mName
End Property

''
' Returns a <b>RegistryKey</b> of the requested SubKey with the write permission specified.
'
' @param Name The name of the SubKey to be opened. Any slashes will will be trimmed from both ends.
' @param Writable The permission to modify the opened SubKey.
' @return A new <b>RegistryKey</b> of the opened SubKey.
' @remarks The requested SubKey is opened starting at the location of the current <b>RegistryKey</b>
' node within the Registry.
' <p>If the requested SubKey was not found, then <b>Nothing</b> is returned. No error is thrown.</p>
'
Public Function OpenSubKey(ByRef Name As String, Optional ByVal Writable As Boolean = False) As RegistryKey
    EnsureNotDisposed
    
    Dim TrimmedName As String
    TrimmedName = CorString.Trim(Name, "\")
    ValidateKey TrimmedName
    
    Dim NewKey As Long
    If RegOpenKeyEx(mHKey, TrimmedName, 0, IIfLong(mWritable, KEY_ALL_ACCESS, KEY_READ), NewKey) = ERROR_SUCCESS Then
        Set OpenSubKey = Cor.NewRegistryKey(NewKey, Path.Combine(mName, TrimmedName), Writable)
    End If
End Function

''
' Closes the <b>RegistryKey</b> object.
'
' @remarks When the key is closed, then it can no longer be manipulated using the <b>RegistryKey</b>
' <p>If the <b>RegistryKey</b> is a system root key, then it will not be closed.</p>
'
Public Sub CloseKey()
    If (Not IsSystemKey) And (mHKey <> vbNullPtr) Then
        RegCloseKey mHKey
        mHKey = vbNullPtr
    End If
End Sub

''
' Flushes any changes to the Windows Registry back to disk.
'
' @remarks Generally this does not need to be called. Windows writes back to the disk Registry
' at regular intervals and when the system shuts down. values of other keys written to the
' Registry may also be flushed back to disk as well.
' <p>Only call this function if it is absolutely necessary to be sure the data has been written
' to the Registry. Excessive calls to this function can impact system performance.</p>
'
Public Sub Flush()
    EnsureNotDisposed
    ThrowOnFailure RegFlushKey(mHKey)
End Sub

''
' Creates a new Windows Registry node.
'
' @param SubKey The name of the new node to be created. Any slashes will be trimmed from the ends.
' @return A <b>RegistryKey</b> to manipulate the new Registry node.
' @remarks If the Registry node already exists, then it is returned.
'
Public Function CreateSubKey(ByRef SubKey As String) As RegistryKey
    EnsureNotDisposed
    VerifyWritable
    
    Dim TrimmedKey As String
    TrimmedKey = CorString.Trim(SubKey, "\")
    ValidateKey TrimmedKey
    
    Dim NewKey As Long
    ThrowOnFailure RegCreateKeyEx(mHKey, TrimmedKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, NewKey, 0)
    
    Set CreateSubKey = Cor.NewRegistryKey(NewKey, Path.Combine(mName, TrimmedKey), True)
End Function


''
' Deletes a node from the Windows Registry.
'
' @param SubKey The node to be deleted.
' @param ThrowOnMissingSubKey Whether to throw an exception if the node was not found.
' @remarks The key node must be empty to be deleted. No sub keys or values are deleted.
' To delete all SubKeys use the <b>DeleteSubKeyTree</b> to recursively delete SubKeys
' and values.
'
Public Sub DeleteSubKey(ByRef SubKey As String, Optional ByVal ThrowOnMissingSubKey As Boolean = True)
    EnsureNotDisposed
    VerifyWritable
    
    Dim Key As RegistryKey
    Set Key = OpenSubKey(SubKey)
    If Not Key Is Nothing Then
        If Key.SubKeyCount > 0 Then _
            Error.InvalidOperation InvalidOperation_RegRemoveSubKey
        
        Key.CloseKey
        ThrowOnFailure RegDeleteKey(mHKey, SubKey)
    ElseIf ThrowOnMissingSubKey Then
        Error.Argument Arg_RegSubKeyAbsent
    End If
End Sub

''
' Deletes all SubKeys within the specified SubKey to delete.
'
' @param SubKey The Registry node to be deleted.
' @remarks All subkeys and values are deleted recursively. Caution should be used.
'
Public Sub DeleteSubKeyTree(ByRef SubKey As String)
    EnsureNotDisposed
    VerifyWritable
    
    Dim TrimmedKey As String
    TrimmedKey = CorString.Trim(SubKey, "\")
    ValidateKey TrimmedKey
    
    Dim Key As RegistryKey
    Set Key = OpenSubKey(TrimmedKey, True)
    
    Dim SubKeys() As String
    SubKeys = Key.GetSubKeyNames
    If CorArray.Length(SubKeys) > 0 Then
        Dim i As Long
        For i = 0 To UBound(SubKeys)
            Key.DeleteSubKeyTree SubKeys(i)
        Next i
    End If
    Key.CloseKey
    
    ThrowOnFailure RegDeleteKey(mHKey, TrimmedKey)
End Sub

''
' Returns a list of SubKey names within the current SubKey.
'
' @return The list of SubKeys.
' @remarks Values within the SubKey are not included in the list.
' <p>If no SubKeys exist then a zero-length array is returned.</p>
'
Public Function GetSubKeyNames() As String()
    EnsureNotDisposed
    
    Dim Name    As String
    Dim Count   As Long
    Dim Names() As String
    
    Name = String$(MaxKeyLength + 1, 0)
    Count = SubKeyCount
    Names = CorArray.CreateInstance(vbString, Count)
    
    Do While Count > 0
        Dim NameLength As Long
        
        Count = Count - 1
        NameLength = MaxKeyLength + 1
        ThrowOnFailure RegEnumKeyEx(mHKey, Count, Name, NameLength, 0, vbNullString, 0, 0@)
        Names(Count) = Left$(Name, NameLength)
    Loop

    GetSubKeyNames = Names
End Function

''
' Returns a list of value names within the current SubKey.
'
' @return The list of value names.
' @remarks SubKeys within the SubKey are not included in the list.
'
' <p>Bad Spec: The MSDN says that if a default value exists, it will
' not be included as part of the names returned. However, a simple
' test in Dot NET shows that the default is included in the list.
' An empty string is the name of the default value. We include
' the default name ("") in the returned list.</p>
'
Public Function GetValueNames() As String()
    EnsureNotDisposed
    
    Dim Name    As String
    Dim Count   As Long
    Dim Names() As String
    
    Name = String$(MaxValueNameLength, 0)
    Count = ValueCount
    Names = CorArray.CreateInstance(vbString, ValueCount)
    
    Do While Count > 0
        Dim NameLength As Long
        
        Count = Count - 1
        NameLength = Len(Name)
        ThrowOnFailure RegEnumValue(mHKey, Count, Name, NameLength, 0, 0, 0, 0)
        Names(Count) = Left$(Name, NameLength)
    Loop
    
    GetValueNames = Names
End Function

''
' Sets the value of a key value within the SubKey.
'
' @param Name The name of the value to set, or an empty string to set the key default value.
' @param Value The value to write to the Registry.
' @param ValueKind The type of value to be written to the registry.
' @remarks The registry can hold several types of data. <b>RegistryKey</b> supports
' many data types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (Expandable string)<br>
' <p>Another value type is any class that implements the <b>cObject</b> interface. The
' <i>ToString</i> method is used to write the string representation to the Registry. This
' makes it easy to write such values as <b>cDateTime</b> with ease.</p>
' <p>If saving as a <b>REG_QWORD</b> then numeric values passed in may be adjusted to
' and mapped to an internal vbCurrency datatype to represent the 64-bit value accurately
' in the registry. If a vbLong value of 1 is passed in, then it will be set to a vbCurrency
' value of 0.0001. This will give the proper bit alignment with the 64-bit registry value.
' In order to avoid having the value mapped to a vbCurrency, pass the value in as a vbCurrency.
' In this case, the value is not changed and the 64-bits are saved to the registry as is.
' A vbCurrency of 1 will be saved to the registry as is, not as 0.0001.</p>
'
Public Sub SetValue(ByRef Name As String, ByRef Value As Variant, Optional ByVal ValueKind As RegistryValueKind = UnknownKind)
    VerifyWritable
    EnsureNotDisposed
    
    If Len(Name) > MaxValueNameLength Then _
        Error.Argument Arg_RegValStrLenBug

    ClearException
    On Error GoTo Catch
    
    Dim Result As Long
    If ValueKind = UnknownKind Then
        Result = SetValueDirect(Name, Value)
    Else
        Result = SetValueWithConvert(Name, Value, ValueKind)
    End If
    On Error GoTo 0
    
    ThrowOnFailure Result
    Exit Sub

Catch:
    Dim Ex As Exception
    If Catch(Ex) Then
        Throw Ex
    Else
        Throw Cor.NewArgumentException("The value did not match the registry type specified by ValueKind.", "value")
    End If
End Sub

''
' Returns the value of a key within a SubKey.
'
' @param Name The name of the value to retrieve, or an empty string to retrieve the key default value.
' @param Default The value to be returned if the key name was not found.
' @param Options Additional options on how to handle how data is handled when reading from the key.
' @return The value found in the Registry, or Empty if the value does not exist.
' @remarks <b>RegistryKey</b> supports the return of 6 different key value types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (String containing an environment variable)<br>
' The specific type will be created and returned based on the type in the registry.
' <p>If retrieving a REG_QWORD, a vbCurrency variable is returned to represent the
' 64-bit value. This means that the value will have the bias of being divided by
' 10000. To get the true 64-bit value, the vbCurrency value will need to be multiplied
' by 10000. The result may be an overflow.
'
Public Function GetValue(ByRef Name As String, Optional ByRef Default As Variant, Optional ByVal Options As RegistryValueOptions) As Variant
    EnsureNotDisposed

    If Len(Name) > MaxValueNameLength Then _
        Error.Argument Arg_RegValStrLenBug

    Dim ValueType   As RegistryValueKind
    Dim ValueSize   As Long
    Dim Result      As Long
    Result = RegQueryValueEx(mHKey, Name, 0, ValueType, ByVal 0, ValueSize)
    
    If Result = ERROR_FILE_NOT_FOUND Then
        If Not IsMissing(Default) Then
            VariantCopyInd GetValue, Default
        End If
        
        Exit Function
    Else
        ThrowOnFailure Result
    End If
    
    Select Case ValueType
        Case StringKind
            Dim sz As String
            sz = String$(ValueSize, 0)
            ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, ByVal StrPtr(sz), LenB(sz))
            GetValue = CorString.TrimEnd(sz, vbNullChar)
            Exit Function
            
        Case DWordKind
            Dim i As Long
            ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, i, 4)
            GetValue = i
            
        Case BinaryKind, REG_NONE
            Dim Bytes() As Byte
            If ValueSize > 0 Then
                ReDim Bytes(0 To ValueSize - 1)
                ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, Bytes(0), ValueSize)
            Else
                Bytes = Cor.NewBytes
            End If
            
            GetValue = Bytes
            Exit Function
            
        Case MultiStringKind
            Dim MultiSZ As String
            MultiSZ = String$(ValueSize, 0)
            ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, ByVal StrPtr(MultiSZ), LenB(MultiSZ))
            GetValue = Split(CorString.TrimEnd(MultiSZ, vbNullChar), vbNullChar)
            Exit Function
            
        Case ExpandStringKind
            Dim ExpSZ As String
            ExpSZ = String$(ValueSize, 0)
            ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, ByVal StrPtr(ExpSZ), LenB(ExpSZ))
            ExpSZ = CorString.TrimEnd(ExpSZ, vbNullChar)
            
            If Options <> DoNotExpandEnvironmentVariables Then
                ExpSZ = Environment.ExpandEnvironmentVariables(ExpSZ)
            End If
            
            GetValue = ExpSZ
            Exit Function
            
        Case QWordKind
            Dim c As Currency
            ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, ValueType, c, 8)
            GetValue = c
            
        Case Else
            VariantCopyInd GetValue, Default
    End Select
End Function

''
' Deletes the value from the registry.
'
' @param Name The name of the value to be deleted, or an empty string to delete the default value for the key.
' @param ThrowOnMissingValue Whether to throw an exception if the value was not found.
'
Public Sub DeleteValue(ByRef Name As String, Optional ByVal ThrowOnMissingValue As Boolean = True)
    VerifyWritable
    EnsureNotDisposed
    
    If (RegDeleteValue(mHKey, Name) = ERROR_FILE_NOT_FOUND) And ThrowOnMissingValue Then
        Error.Argument Arg_RegSubKeyValueAbsent
    End If
End Sub

''
' Returns the number of SubKeys within the current key.
'
' @return The number of SubKeys.
'
Public Property Get SubKeyCount() As Long
    EnsureNotDisposed
    ThrowOnFailure RegQueryInfoKey(mHKey, vbNullString, ByVal 0, 0, SubKeyCount, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0)
End Property

''
' Returns the number of values within the current key.
'
' @remarks The number of values.
'
Public Property Get ValueCount() As Long
    EnsureNotDisposed
    ThrowOnFailure RegQueryInfoKey(mHKey, vbNullString, ByVal 0, 0, ByVal 0, ByVal 0, ByVal 0, ValueCount, ByVal 0, ByVal 0, ByVal 0, ByVal 0)
End Property

''
' Returns the type of value stored in the registry.
'
' @param Name The name of the value to get the type of.
' @return The type of value in the registry.
'
Public Function GetValueKind(ByRef Name As String) As RegistryValueKind
    EnsureNotDisposed
    ThrowOnFailure RegQueryValueEx(mHKey, Name, 0, GetValueKind, ByVal 0, 0)
End Function

''
' Returns the last time the subkey was written to.
'
' @return A CorDateTime object.
' @remarks This method returns CorDateTime.MinValue on Windows 95,98 and ME.
'
Public Function GetLastWriteTime() As CorDateTime
    EnsureNotDisposed
    
    If Environment.IsNT Then
        Dim Time As Currency
        ThrowOnFailure RegQueryInfoKey(mHKey, vbNullString, ByVal 0, 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, Time)
        Set GetLastWriteTime = CorDateTime.FromFileTime(Time)
    Else
        Set GetLastWriteTime = CorDateTime.MinValue
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Key As Long, ByRef Name As String, ByVal Writable As Boolean)
    mHKey = Key
    mName = CorString.Trim(Name, "\")
    mWritable = Writable
End Sub

Private Sub Class_Terminate()
    CloseKey
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ThrowOnFailure(ByVal Result As Long)
    If Result <> ERROR_SUCCESS Then
        Error.Win32Error Result
    End If
End Sub

Private Sub ValidateKey(ByRef Name As String)
    Dim NextSlash   As Long
    Dim Current     As Long
    
    NextSlash = InStr(Name, "\") - 1
    Do While NextSlash > 0
        If NextSlash - Current > MaxKeyLength Then
            Error.Argument Arg_RegKeyStrLenBug
        End If
        
        Current = NextSlash + 1
        NextSlash = InStr(Current + 1, Name, "\") - 1
    Loop
    
    If Len(Name) - Current > MaxKeyLength Then
        Error.Argument Arg_RegKeyStrLenBug
    End If
End Sub

Private Function IsSystemKey() As Boolean
    IsSystemKey = (mHKey >= &H80000005 And mHKey <= &H80000000)
End Function

Private Sub EnsureNotDisposed()
    If mHKey = vbNullPtr Then
        Error.ObjectDisposed ObjectDisposed_RegKeyClosed
    End If
End Sub

Private Sub VerifyWritable()
    If Not mWritable Then
        Error.UnauthorizedAccess UnauthorizedAccess_RegistryNoWrite
    End If
End Sub

Private Function SetValueSZ(ByRef Name As String, ByRef Value As Variant) As Long
    Dim Data As String
    Data = Value & vbNullChar
    SetValueSZ = RegSetValueEx(mHKey, Name, 0, REG_SZ, StrPtr(Data), LenB(Data))
End Function

Private Function SetValueMultiSZ(ByRef Name As String, ByRef Value As Variant) As Long
    Dim s As String
    If CorArray.IsNull(Value) Then
        s = vbNullChar
    ElseIf CorArray.Length(Value) > 0 Then
        s = Join(Value, vbNullChar) & (vbNullChar & vbNullChar)
    Else
        s = vbNullChar
    End If
    
    SetValueMultiSZ = RegSetValueEx(mHKey, Name, 0, REG_MULTI_SZ, StrPtr(s), LenB(s))
End Function

Private Function SetValueExpandSZ(ByRef Name As String, ByRef Value As Variant) As Long
    Dim Data As String
    Data = Value & vbNullChar
    SetValueExpandSZ = RegSetValueEx(mHKey, Name, 0, REG_EXPAND_SZ, StrPtr(Data), LenB(Data))
End Function

Private Function SetValueDWord(ByRef Name As String, ByVal Value As Long) As Long
    SetValueDWord = RegSetValueEx(mHKey, Name, 0, REG_DWORD, VarPtr(Value), 4)
End Function

Private Function SetValueBinary(ByRef Name As String, ByRef Value As Variant) As Long
    Dim lpData  As Long
    Dim Size    As Long
    
    If Not CorArray.IsNull(Value) Then
        Size = CorArray.Length(Value)
        If Size > 0 Then
            lpData = MemLong(CorArray.ArrayPointer(Value) + PVDATA_OFFSET)
        End If
    End If
    
    SetValueBinary = RegSetValueEx(mHKey, Name, 0, REG_BINARY, lpData, Size)
End Function

Private Function SetValueQWord(ByRef Name As String, ByVal Value As Currency) As Long
    SetValueQWord = RegSetValueEx(mHKey, Name, 0, REG_QWORD, VarPtr(Value), 8)
End Function

Private Function SetValueObject(ByRef Name As String, ByVal Value As IObject) As Long
    SetValueObject = SetValueSZ(Name, Value.ToString)
End Function

Private Function SetValueDirect(ByRef Name As String, ByRef Value As Variant)
    Dim Result As Long
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            Result = SetValueDWord(Name, Value)
        Case corStringArray, corVariantArray
            Result = SetValueMultiSZ(Name, Value)
        Case corByteArray
            Result = SetValueBinary(Name, Value)
        Case vbCurrency
            Result = SetValueQWord(Name, Value)
        Case vbObject
            Result = SetValueObject(Name, Value)
        Case Else
            Result = SetValueSZ(Name, Value)
    End Select
    
    SetValueDirect = Result
End Function

Private Function SetValueWithConvert(ByRef Name As String, ByRef Value As Variant, ByVal ValueKind As RegistryValueKind) As Long
    Dim Result As Long
    
    Select Case ValueKind
        Case DWordKind
            Result = SetValueDWord(Name, Value)
        Case BinaryKind
            Result = SetValueBinary(Name, Value)
        Case MultiStringKind
            Result = SetValueMultiSZ(Name, Value)
        Case ExpandStringKind
            Result = SetValueExpandSZ(Name, Value)
        Case QWordKind
            ' We assume if an actual Currency datatype value was passed in
            ' then that value was the expected 64-bit value. If a different
            ' datatype (eg. vbLong) was passed in, then we assume that the
            ' 64-bits need to be lined up. A vbLong of 1 does not equal
            ' a vbCurrency of 1 in a 64-bit representation, so we need to
            ' divide the vbLong value by 10000 to have correct bit alignment.
            If VarType(Value) = vbCurrency Then
                Result = SetValueQWord(Name, Value)
            Else
                Result = SetValueQWord(Name, 0.0001@ * CCur(Value))
            End If
            
        Case Else
            If IsObject(Value) Then
                Result = SetValueObject(Name, Value)
            Else
                Result = SetValueSZ(Name, Value)
            End If
    End Select
    
    SetValueWithConvert = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
