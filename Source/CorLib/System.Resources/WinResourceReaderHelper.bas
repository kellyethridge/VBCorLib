Attribute VB_Name = "WinResourceReaderHelper"
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: modWinResourceReader
'

Option Explicit

Public Function EnumResTypeProc(ByVal hModule As Long, ByVal lpszType As Long, ByRef Reader As WinResourceReader) As Long
    EnumResTypeProc = EnumResourceNames(hModule, lpszType, AddressOf EnumResNameProc, VarPtr(Reader))
End Function

Private Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByRef Reader As WinResourceReader) As Long
    EnumResNameProc = EnumResourceLanguages(hModule, lpszType, lpszName, AddressOf EnumResLangProc, VarPtr(Reader))
End Function

Private Function EnumResLangProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal wIDLanguage As Integer, ByRef Reader As WinResourceReader) As Long
    Dim SearchHandle As Long
    SearchHandle = FindResourceEx(hModule, lpszType, lpszName, wIDLanguage)
    
    If SearchHandle <> NULL_HANDLE Then
        Dim ResourceHandle As Long
        ResourceHandle = LoadResource(hModule, SearchHandle)
        If ResourceHandle <> NULL_HANDLE Then
            Dim lpData As Long
            lpData = LockResource(ResourceHandle)
            
            Dim ResInBytes As Long
            ResInBytes = SizeofResource(hModule, SearchHandle)
            
            ' Copy our raw byte data from the resource.
            Dim Data() As Byte
            ReDim Data(0 To ResInBytes - 1)
            CopyMemory Data(0), ByVal lpData, ResInBytes
            
            ' Get the resource type, either a string or number.
            Dim ResourceType As Variant
            ResourceType = GetOrdinalOrName(lpszType)
            
            ' Get the resource name, either a string or a number.
            Dim ResourceName As Variant
            ResourceName = GetOrdinalOrName(lpszName)
            
            Dim Key As ResourceKey
            Set Key = Cor.NewResourceKey(ResourceName, ResourceType, wIDLanguage)
            
            Reader.AddResource Key, Data
        End If
    End If
    
    EnumResLangProc = BOOL_TRUE
End Function

Private Function GetOrdinalOrName(ByVal Ptr As Long) As Variant
    If Ptr And &HFFFF0000 Then
        ' we have a string and Ptr is
        ' the pointer to the first character.
        
        ' Finds the length by finding the null character.
        Dim StringLength As Long
        StringLength = lstrlen(Ptr)
        
        Dim Chars() As Byte
        ReDim Chars(0 To StringLength - 1)
        
        CopyMemory Chars(0), ByVal Ptr, StringLength
        GetOrdinalOrName = StrConv(Chars, vbUnicode)
    Else
        ' we have a number, so just return it.
        GetOrdinalOrName = Ptr
    End If
End Function

