Attribute VB_Name = "CharAllocation"
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
' Module: CharAllocation
'

''
' Provides a central location to create Integer array proxy access to Strings.
'
' @remarks <p>A proxy char buffer is used with the backing of the string value
' passed in. Once the buffer access is no longer needed then the FreeChars
' method is called, passing in the Integer array returned during allocation.</p>
' <p>A Variant that contains either a String or Integer array can also be accessed
' as an array by calling AsChars. If the Variant contains a String type then the
' process works the same as calling AllocChars. If the Variant contains an Integer
' array, then the array itself is returned without allocating any char buffers.
' Which ever method is used, FreeChars must still be called using the original
' array returned to remove having multiple handles point to the same data.
'
Option Explicit

Private Const BufferCapacity As Long = 16

Private Type BufferBucket
    TablePtr    As Long
    Self        As IUnknown
    ReleasePtr  As Long
    Buffer      As SafeArray1d
    BufferPtr   As Long
    InUse       As Boolean
End Type

Public Type CharBuffer
    TablePtr    As Long
    Self        As IUnknown
    ReleasePtr  As Long
    Chars()     As Integer
    Buffer      As SafeArray1d
End Type

Private mInited         As Boolean
Private mBuckets()      As BufferBucket
Private mFastLaneBucket As BufferBucket


Public Sub InitChars(ByRef Buffer As CharBuffer, Optional ByRef s As String)
    With Buffer
        .TablePtr = VarPtr(.TablePtr)
        ObjectPtr(.Self) = .TablePtr
        .ReleasePtr = FuncAddr(AddressOf ReleaseCharBuffer)
        SAPtr(.Chars) = VarPtr(.Buffer)
        
        With .Buffer
            .cbElements = vbSizeOfChar
            .cDims = 1
            .cLocks = 1
            .pvData = StrPtr(s)
            .cElements = Len(s)
        End With
    End With
End Sub

Public Sub SetChars(ByRef Buffer As CharBuffer, ByRef s As String)
    With Buffer.Buffer
        .pvData = StrPtr(s)
        .cElements = Len(s)
    End With
End Sub

''
' Allocates an Integer array backed by the String passed in.
'
' Once work is finished with the array, FreeChars must be called to remove
' any references to the original string value.
'
Public Function AllocChars(ByRef s As String) As Integer()
    Dim Index As Long
    
    ' >99% of the time only a single allocation will be in effect,
    ' so create a fastlane to improve efficiency by removing the
    ' need to call into a function and perform a look-up for an
    ' empty bucket. This goes from an O(n) to O(1) efficiency.
    If Not mFastLaneBucket.InUse Then
        If mFastLaneBucket.BufferPtr = vbNullPtr Then
            InitBucket mFastLaneBucket
        End If
        
        FillBucket mFastLaneBucket, s
        SAPtr(AllocChars) = mFastLaneBucket.BufferPtr
    Else
        Index = FindAvailableBucketIndex
        FillBucket mBuckets(Index), s
        SAPtr(AllocChars) = mBuckets(Index).BufferPtr
    End If
End Function

''
' Either allocs or returns an Integer array.
'
' @param v A Variant containing either a String or Integer() data type.
' @return Returns a reference to an Integer array to access the original value.
' @remarks <p>If a String is passed in then the AllocChars method is used to create
' an Integer array with the string as a backing-store. If an Integer array is
' passed in, then another reference to the array is returned.</p>
' <p>Once work is finished with the array, FreeChars must be called to remove
' any references to the original string or array.</p>
'
Public Function AsChars(ByRef v As Variant) As Integer()
    Select Case VarType(v)
        Case vbString
            ' Directly assigning a string pointer prevents a string from being copied.
            Dim LocalString As String
            StringPtr(LocalString) = StrPtr(v)
            AsChars = AllocChars(LocalString)
            StringPtr(LocalString) = vbNullPtr
            
        Case vbIntegerArray
            ' Directly assigning an array pointer prevents the source array from being copied.
            SAPtr(AsChars) = SAPtrV(v)
            
        Case Else
            Error.Argument Argument_CharArrayRequired
            
    End Select
End Function

''
' Removes the refernece the allocated Integer array points to.
'
' @param Chars The array that the reference is removed from.
' @remarks If the original backing of the array was a string, then
' the internal char-buffer is cleared as well.
'
Public Sub FreeChars(ByRef Chars() As Integer)
    Dim Index As Long
    
    If SAPtr(Chars) = mFastLaneBucket.BufferPtr Then
        mFastLaneBucket.InUse = False
    Else
        Index = FindAllocatedBucketIndex(Chars)
        
        If Index >= 0 Then
            mBuckets(Index).InUse = False
        End If
    End If
    
    SAPtr(Chars) = vbNullPtr
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitBuckets()
    ReDim mBuckets(0 To BufferCapacity - 1)

    Dim i As Long
    For i = 0 To UBound(mBuckets)
        InitBucket mBuckets(i)
    Next
End Sub

Private Sub InitBucket(ByRef Bucket As BufferBucket)
    With Bucket
        .Buffer.cbElements = 2
        .Buffer.cDims = 1
        .Buffer.cLocks = 1
        .TablePtr = VarPtr(.TablePtr)
        ObjectPtr(.Self) = .TablePtr
        .ReleasePtr = FuncAddr(AddressOf ReleaseBufferBucket)
        .BufferPtr = VarPtr(.Buffer)
    End With
End Sub

Private Sub FillBucket(ByRef Bucket As BufferBucket, ByRef s As String)
    With Bucket
        .InUse = True
        
        With .Buffer
            .cElements = Len(s)
            .pvData = StrPtr(s)
        End With
    End With
End Sub

Private Function FindAvailableBucketIndex() As Long
    If Not mInited Then
        InitBuckets
        mInited = True
    End If

    Dim i As Long
    For i = 0 To BufferCapacity - 1
        If Not mBuckets(i).InUse Then
            FindAvailableBucketIndex = i
            Exit Function
        End If
    Next
    
    Debug.Assert False
End Function

Private Function FindAllocatedBucketIndex(ByRef Chars() As Integer) As Long
    Dim Ptr As Long
    Dim i   As Long
    
    If mInited Then
        Ptr = SAPtr(Chars)
        
        For i = 0 To BufferCapacity - 1
            If mBuckets(i).BufferPtr = Ptr Then
                FindAllocatedBucketIndex = i
                Exit Function
            End If
        Next
    End If
    
    FindAllocatedBucketIndex = -1
End Function

Private Function ReleaseBufferBucket(ByRef This As BufferBucket) As Long
    This.Buffer.pvData = vbNullPtr
    This.Buffer.cElements = 0
    This.Buffer.cLocks = 0
End Function

Private Function ReleaseCharBuffer(ByRef This As CharBuffer) As Long
    This.Buffer.pvData = vbNullPtr
    This.Buffer.cElements = 0
    This.Buffer.cLocks = 0
    SAPtr(This.Chars) = vbNullPtr
End Function
