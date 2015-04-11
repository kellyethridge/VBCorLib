Attribute VB_Name = "DataTypeConstants"
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: DataTypeConstants
'
Option Explicit

Public Const vbIntegerArray         As Long = vbInteger Or vbArray
Public Const vbByteArray            As Long = vbByte Or vbArray
Public Const vbLongArray            As Long = vbLong Or vbArray
Public Const vbBooleanArray         As Long = vbBoolean Or vbArray
Public Const vbStringArray          As Long = vbString Or vbArray
Public Const vbVariantArray         As Long = vbVariant Or vbArray

Public Const SizeOfByte             As Long = 1
Public Const SizeOfInteger          As Long = 2
Public Const SizeOfLong             As Long = 4
Public Const SizeOfSingle           As Long = 4
Public Const SizeOfDouble           As Long = 8
Public Const SizeOfCurrency         As Long = 8
Public Const SizeOfDecimal          As Long = 16
Public Const SizeOfBoolean          As Long = 1
Public Const SizeOfDate             As Long = 8
Public Const SizeOfSafeArray        As Long = 16
Public Const SizeOfSafeArrayBound   As Long = 8
Public Const SizeOfSafeArray1d      As Long = SizeOfSafeArray + SizeOfSafeArrayBound
Public Const SizeOfGuid             As Long = 16
Public Const SizeOfGuidSafeArray1d  As Long = SizeOfSafeArray1d + SizeOfGuid
Public Const SizeOfVariant          As Long = 16

' Byte offsets into the SafeArray structure.
Public Const FFEATURES_OFFSET               As Long = 2
Public Const CBELEMENTS_OFFSET              As Long = 4
Public Const PVDATA_OFFSET                  As Long = 12
Public Const LBOUND_OFFSET                  As Long = 20
Public Const CLOCKS_OFFSET                  As Long = 8
Public Const CELEMENTS_OFFSET               As Long = 16

' Variant descriptions and offsets into the layout.
Public Const VARIANTDATA_OFFSET             As Long = 8
Public Const VT_BYREF                       As Long = &H4000
