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
Implements IEnumerable


Private Const REG_OPTION_NON_VOLATILE       As Long = 0
Private Const MAX_KEY_LENGTH                As Long = 255
Private Const DEFAULT_VALUE_NAME            As String = ""


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
    Call VerifyHandle
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
Public Function OpenSubKey(ByVal Name As String, Optional ByVal Writable As Boolean = False) As RegistryKey
    Call VerifyHandle
    
    Name = cString.Trim(Name, "\")
    Call VerifyKey(Name)
    
    Dim Access As Long
    If mWritable Then
        Access = KEY_ALL_ACCESS
    Else
        Access = KEY_READ
    End If
    
    Dim Result As Long
    Dim NewKey As Long
    Result = Api.RegOpenKeyEx(mHKey, Name, 0, Access, NewKey)
    If Result <> ERROR_SUCCESS Then Exit Function

    Set OpenSubKey = Cor.NewRegistryKey(NewKey, Path.Combine(mName, Name), Writable)
End Function

''
' Closes the <b>RegistryKey</b> object.
'
' @remarks When the key is closed, then it can no longer be manipulated using the <b>RegistryKey</b>
' <p>If the <b>RegistryKey</b> is a system root key, then it will not be closed.</p>
'
Public Sub CloseKey()
    If (Not IsSystemKey) And (mHKey <> vbNullPtr) Then
        Call RegCloseKey(mHKey)
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
    Call VerifyHandle
    
    Dim Result As Long
    Result = RegFlushKey(mHKey)
    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey.Flush"
End Sub

''
' Creates a new Windows Registry node.
'
' @param SubKey The name of the new node to be created. Any slashes will be trimmed from the ends.
' @return A <b>RegistryKey</b> to manipulate the new Registry node.
' @remarks If the Registry node already exists, then it is returned.
'
Public Function CreateSubKey(ByVal SubKey As String) As RegistryKey
    Call VerifyHandle
    Call VerifyWritable
    
    SubKey = cString.Trim(SubKey, "\")
    Call VerifyKey(SubKey)
    
    Dim Result As Long
    Dim NewKey As Long
    Result = Api.RegCreateKeyEx(mHKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, NewKey, 0)
    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
    
    Set CreateSubKey = Cor.NewRegistryKey(NewKey, Path.Combine(mName, SubKey), True)
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
Public Sub DeleteSubKey(ByVal SubKey As String, Optional ByVal ThrowOnMissingSubKey As Boolean = True)
    Call VerifyHandle
    Call VerifyWritable
    
    Dim rk As RegistryKey
    Set rk = Me.OpenSubKey(SubKey)
    
    If rk Is Nothing Then
        If ThrowOnMissingSubKey Then _
            Throw Cor.NewArgumentException("The specified SubKey was not found.", "SubKey")
    Else
        If rk.SubKeyCount > 0 Then _
            Throw Cor.NewInvalidOperationException("Cannot delete. SubKey contains keys.")
        
        Dim Result As Long
        Result = Api.RegDeleteKey(mHKey, SubKey)
        If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
        rk.CloseKey
    End If
End Sub

''
' Deletes all SubKeys within the specified SubKey to delete.
'
' @param SubKey The Registry node to be deleted.
' @remarks All subkeys and values are deleted recursively. Caution should be used.
'
Public Sub DeleteSubKeyTree(ByVal SubKey As String)
    Call VerifyHandle
    Call VerifyWritable
    
    SubKey = cString.Trim(SubKey, "\")
    Call VerifyKey(SubKey)
    
    Dim rk As RegistryKey
    Set rk = Me.OpenSubKey(SubKey, True)
    
    Dim SubKeys() As String
    SubKeys = rk.GetSubKeyNames
    If CorArray.GetLength(SubKeys) > 0 Then
        Dim i As Long
        For i = 0 To UBound(SubKeys)
            Call rk.DeleteSubKeyTree(SubKeys(i))
        Next i
    End If
    Call rk.CloseKey
    
    Dim Result As Long
    Result = Api.RegDeleteKey(mHKey, SubKey)
    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey.DeleteSubKeyTree"
End Sub

''
' Returns a list of SubKey names within the current SubKey.
'
' @return The list of SubKeys.
' @remarks Values within the SubKey are not included in the list.
' <p>If no SubKeys exist then a zero-length array is returned.</p>
'
Public Function GetSubKeyNames() As String()
    Call VerifyHandle
    
    Dim Buf As String
    Buf = String$(MAX_KEY_LENGTH + 1, 0)
    
    Dim KeyCount As Long
    KeyCount = Me.SubKeyCount
    
    Dim Ret() As String
    Ret = CorArray.CreateInstance(vbString, KeyCount)
    
    Do While KeyCount > 0
        KeyCount = KeyCount - 1
        
        Dim BufferSize As Long
        BufferSize = MAX_KEY_LENGTH + 1
        
        Dim Result As Long
        Result = Api.RegEnumKeyEx(mHKey, KeyCount, Buf, BufferSize, 0, vbNullString, 0, 0@)
        If Result <> ERROR_SUCCESS Then IOError Result, "GetSubKeyNames"
        
        Ret(KeyCount) = Left$(Buf, BufferSize)
    Loop

    GetSubKeyNames = Ret
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
    Call VerifyHandle
    
    Dim Buf As String
    Buf = String$(Registry.MaxValueNameLength, 0)
    
    Dim ValueCount As Long
    ValueCount = Me.ValueCount
    
    Dim Ret() As String
    Ret = CorArray.CreateInstance(vbString, ValueCount)
    
    Do While ValueCount > 0
        ValueCount = ValueCount - 1
        
        Dim BufferSize As Long
        BufferSize = Len(Buf)
        
        Dim Result As Long
        Result = Api.RegEnumValue(mHKey, ValueCount, Buf, BufferSize, 0, 0, 0, 0)
        If Result <> ERROR_SUCCESS Then IOError Result, "GetValueNames"
        
        Ret(ValueCount) = Left$(Buf, BufferSize)
    Loop
    
    GetValueNames = Ret
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
Public Sub SetValue(ByVal Name As String, ByRef Value As Variant, Optional ByVal ValueKind As RegistryValueKind = UnknownKind)
    Call VerifyWritable
    Call VerifyHandle
    
    If Len(Name) > Registry.MaxValueNameLength Then _
        Throw Cor.NewArgumentException("Registry Value name must not exceed " & Registry.MaxValueNameLength & " characters.", "Name")

    Call ClearException
    On Error GoTo errTrap
    
    Dim Result As Long
    If ValueKind = UnknownKind Then
        Result = SetValueDirect(Name, Value)
    Else
        Result = SetValueWithConvert(Name, Value, ValueKind)
    End If
    On Error GoTo 0
    
    If Result <> ERROR_SUCCESS Then IOError Result
    Exit Sub

errTrap:
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
Public Function GetValue(ByVal Name As String, Optional ByRef Default As Variant, Optional ByVal Options As RegistryValueOptions) As Variant
    Call VerifyHandle

    If Len(Name) > Registry.MaxValueNameLength Then _
        Throw Cor.NewArgumentException("Registry Value name must not exceed " & Registry.MaxValueNameLength & " characters.", "Name")

    Dim ValueType   As RegistryValueKind
    Dim ValueSize   As Long
    Dim Result      As Long
    Result = Api.RegQueryValueEx(mHKey, Name, ValueType, 0, ValueSize)
    
    If Result = ERROR_FILE_NOT_FOUND Then
        If Not IsMissing(Default) Then Call VariantCopyInd(GetValue, Default)
        Exit Function
    ElseIf Result <> ERROR_SUCCESS Then
        IOError Result
    End If
    
    Select Case ValueType
        Case StringKind
            Dim sz As String
            sz = String$(ValueSize, 0)
            Result = Api.RegQueryValueExStr(mHKey, Name, ValueType, sz)
            If Result <> ERROR_SUCCESS Then IOError Result
            GetValue = cString.TrimEnd(sz, vbNullChar)
            Exit Function
            
        Case DWordKind
            Dim i As Long
            Result = Api.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(i), 4)
            If Result <> ERROR_SUCCESS Then IOError Result
            GetValue = i
            
        Case BinaryKind, REG_NONE
            Dim Bytes() As Byte
            If ValueSize > 0 Then
                ReDim Bytes(0 To ValueSize - 1)
                Result = Api.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(Bytes(0)), ValueSize)
                If Result <> ERROR_SUCCESS Then IOError Result
            Else
                Bytes = Cor.NewBytes
            End If
            
            GetValue = Bytes
            Exit Function
            
        Case MultiStringKind
            Dim MultiSZ As String
            MultiSZ = String$(ValueSize, 0)
            Result = Api.RegQueryValueExStr(mHKey, Name, ValueType, MultiSZ)
            If Result <> ERROR_SUCCESS Then IOError Result
            GetValue = Split(cString.TrimEnd(MultiSZ, vbNullChar), vbNullChar)
            Exit Function
            
        Case ExpandStringKind
            Dim ExpSZ As String
            ExpSZ = String$(ValueSize, 0)
            Result = Api.RegQueryValueExStr(mHKey, Name, ValueType, ExpSZ)
            If Result <> ERROR_SUCCESS Then IOError Result
            ExpSZ = cString.TrimEnd(ExpSZ, vbNullChar)
            
            If Options <> DoNotExpandEnvironmentVariables Then
                ExpSZ = Environment.ExpandEnvironmentVariables(ExpSZ)
            End If
            
            GetValue = ExpSZ
            Exit Function
            
        Case QWordKind
            Dim c As Currency
            Result = Api.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(c), 8)
            If Result <> ERROR_SUCCESS Then IOError Result
            GetValue = c
            
        Case Else
            Throw Cor.NewArgumentException("Unsupported Registry type.")
    End Select
End Function

''
' Deletes the value from the registry.
'
' @param Name The name of the value to be deleted, or an empty string to delete the default value for the key.
' @param ThrowOnMissingValue Whether to throw an exception if the value was not found.
'
Public Sub DeleteValue(ByVal Name As String, Optional ByVal ThrowOnMissingValue As Boolean = True)
    Call VerifyWritable
    Call VerifyHandle
    
    Dim Result As Long
    Result = Api.RegDeleteValue(mHKey, Name)
    If (Result = ERROR_FILE_NOT_FOUND) And ThrowOnMissingValue Then _
        Throw Cor.NewArgumentException("Value was not found.", "Name")
End Sub

''
' Returns the number of SubKeys within the current key.
'
' @return The number of SubKeys.
'
Public Property Get SubKeyCount() As Long
    Call VerifyHandle
    
    Dim Result As Long
    Result = Api.RegQueryInfoKey(mHKey, vbNullString, 0, 0, VarPtr(SubKeyCount), 0, 0, 0, 0, 0, 0, 0)
    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
End Property

''
' Returns the number of values within the current key.
'
' @remarks The number of values.
'
Public Property Get ValueCount() As Long
    Call VerifyHandle
    
    Dim Result As Long
    Result = Api.RegQueryInfoKey(mHKey, vbNullString, 0, 0, 0, 0, 0, VarPtr(ValueCount), 0, 0, 0, 0)
    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
End Property

''
' Returns the type of value stored in the registry.
'
' @param Name The name of the value to get the type of.
' @return The type of value in the registry.
'
Public Function GetValueKind(ByVal Name As String) As RegistryValueKind
    Call VerifyHandle
    
    Dim Result As Long
    Result = Api.RegQueryValueEx(mHKey, Name, GetValueKind, 0, 0)
    If Result <> ERROR_SUCCESS Then IOError Result, "GetValueKind"
End Function

''
' Returns the last time the subkey was written to.
'
' @return A cDateTime object.
' @remarks This method returns cDateTime.MinValue on Windows 95,98 and ME.
'
Public Function GetLastWriteTime() As cDateTime
    Call VerifyHandle
    
    If Environment.IsNT Then
        Dim Result  As Long
        Dim Time    As Currency
        
        Result = Api.RegQueryInfoKey(mHKey, vbNullString, 0, 0, 0, 0, 0, 0, 0, 0, 0, VarPtr(Time))
        If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
        
        Set GetLastWriteTime = cDateTime.FromFileTime(Time)
    Else
        Set GetLastWriteTime = cDateTime.MinValue
    End If
End Function

''
' Returns an enumerator to enumerate the name/value pairs.
'
' @param EnumType A flag indicating which type of values to be enumerated in the key.
' @param Options Additional options on how to handle registry values.
' @return An IEnumerator object.
'
Public Function GetEnumerator(Optional ByVal EnumType As IDictionaryEnumType = detEntries, Optional ByVal Options As RegistryValueOptions) As Object
    Call VerifyHandle
    
    Dim Ret As New RegistryKeyEnumerator
    Call Ret.Init(Me, EnumType, Options)
    Set GetEnumerator = Ret
End Function

''
' Returns a For..Each compatible enumerator.
'
' @return A For..Each compatible enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Key As Long, ByVal Name As String, ByVal Writable As Boolean)
    mHKey = Key
    mName = cString.Trim(Name, "\")
    mWritable = Writable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyKey(ByRef Path As String)
    Dim Keys() As String
    Keys = Split(Path, "\")
    
    Dim i As Long
    For i = 0 To UBound(Keys)
        If Len(Keys(i)) > MAX_KEY_LENGTH Then
            Throw Cor.NewArgumentException("Key names cannot exceed 255 characters.")
        End If
    Next i
End Sub

Private Function IsSystemKey() As Boolean
    If mHKey < &H80000006 Then Exit Function
    If mHKey > &H80000000 Then Exit Function
    IsSystemKey = True
End Function

Private Sub VerifyHandle()
    If mHKey = vbNullPtr Then _
        Throw Cor.NewObjectDisposedException("RegistryKey", "The Registry Key is closed.")
    
    If mHKey = HKEY_DYN_DATA Then
        If Environment.IsNT Then IOError ERROR_INVALID_HANDLE
    End If
End Sub

Private Sub VerifyWritable()
    If Not mWritable Then Throw Cor.NewUnauthorizedAccessException("The Registry Key is not writable.")
End Sub

Private Function SetValueSZ(ByRef Name As String, ByRef Value As Variant) As Long
    SetValueSZ = Api.RegSetValueExStr(mHKey, Name, REG_SZ, Value & vbNullChar)
End Function

Private Function SetValueMultiSZ(ByRef Name As String, ByRef Value As Variant) As Long
    Dim s As String
    If CorArray.IsNull(Value) Then
        s = vbNullChar
    ElseIf CorArray.GetLength(Value) > 0 Then
        s = Join(Value, vbNullChar) & (vbNullChar & vbNullChar)
    Else
        s = vbNullChar
    End If
    
    SetValueMultiSZ = Api.RegSetValueExStr(mHKey, Name, REG_MULTI_SZ, s)
End Function

Private Function SetValueExpandSZ(ByRef Name As String, ByRef Value As Variant) As Long
    SetValueExpandSZ = Api.RegSetValueExStr(mHKey, Name, REG_EXPAND_SZ, Value & vbNullChar)
End Function

Private Function SetValueDWord(ByRef Name As String, ByVal Value As Long) As Long
    SetValueDWord = Api.RegSetValueEx(mHKey, Name, REG_DWORD, VarPtr(Value), 4)
End Function

Private Function SetValueBinary(ByRef Name As String, ByRef Value As Variant) As Long
    Dim lpData  As Long
    Dim Size    As Long
    
    If Not CorArray.IsNull(Value) Then
        Size = CorArray.GetLength(Value)
        If Size > 0 Then lpData = MemLong(CorArray.ArrayPointer(Value) + PVDATA_OFFSET)
    End If
    
    SetValueBinary = Api.RegSetValueEx(mHKey, Name, REG_BINARY, lpData, Size)
End Function

Private Function SetValueQWord(ByRef Name As String, ByVal Value As Currency) As Long
    SetValueQWord = Api.RegSetValueEx(mHKey, Name, REG_QWORD, VarPtr(Value), 8)
End Function

Private Function SetValueObject(ByRef Name As String, ByVal Value As IObject) As Long
    SetValueObject = SetValueSZ(Name, Value.ToString)
End Function

Private Function SetValueDirect(ByRef Name As String, ByRef Value As Variant)
    Dim Result As Long
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte:     Result = SetValueDWord(Name, Value)
        Case vbStringArray, vbVariantArray: Result = SetValueMultiSZ(Name, Value)
        Case vbByteArray:                   Result = SetValueBinary(Name, Value)
        Case vbCurrency:                    Result = SetValueQWord(Name, Value)
        Case vbObject:                      Result = SetValueObject(Name, Value)
        Case Else:                          Result = SetValueSZ(Name, Value)
    End Select
    
    SetValueDirect = Result
End Function

Private Function SetValueWithConvert(ByRef Name As String, ByRef Value As Variant, ByVal ValueKind As RegistryValueKind) As Long
    Dim Result As Long
    
    Select Case ValueKind
        Case DWordKind:         Result = SetValueDWord(Name, Value)
        Case BinaryKind:        Result = SetValueBinary(Name, Value)
        Case MultiStringKind:   Result = SetValueMultiSZ(Name, Value)
        Case ExpandStringKind:  Result = SetValueExpandSZ(Name, Value)
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
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call CloseKey
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
