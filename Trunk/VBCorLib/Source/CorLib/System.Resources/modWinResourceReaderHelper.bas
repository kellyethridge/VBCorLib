Attribute VB_Name = "modWinResourceReader"
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modWinResourceReader
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
            Call CopyMemory(Data(0), ByVal lpData, ResInBytes)
            
            ' Get the resource type, either a string or number.
            Dim ResourceType As Variant
            ResourceType = GetOrdinalOrName(lpszType)
            
            ' Get the resource name, either a string or a number.
            Dim ResourceName As Variant
            ResourceName = GetOrdinalOrName(lpszName)
            
            Dim Key As ResourceKey
            Set Key = Cor.NewResourceKey(ResourceName, ResourceType, wIDLanguage)
            
            Call Reader.AddResource(Key, Data)
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
        
        Call CopyMemory(Chars(0), ByVal Ptr, StringLength)
        GetOrdinalOrName = StrConv(Chars, vbUnicode)
    
    Else
        ' we have a number, so just return it.
        GetOrdinalOrName = Ptr
    End If
End Function

