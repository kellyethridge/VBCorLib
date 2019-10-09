VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitMapResourceDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: BitmapResourceDecoder
'

''
' Decodes bitmap byte data from a byte array and produces a StdPicture.
' Usually the byte array comes from a .RES file file.
'
' @see ResourceReader
' @see WinResourceReader
'
Option Explicit
Implements IResourceDecoder
Private Const SIZEOF_BITMAPFILEHEADER As Long = 14
Private Const GMEM_MOVEABLE As Long = &H2

Private mKey    As ResourceKey
Private mBitmap As StdPicture


''
' Attempts to decode a set of bytes into a Bitmap.
'
' @param Key The identifier for the set of bytes.
' @param Bytes The byte data to be decoded into a bitmap.
' @return The number of resources decoded from the byte data. Returns 1 on success.
' @remarks The key must be of a resource type bitmap (ResourceTypes.BitmapResource).
' @see ResourceTypes
' @see ResourceKey
'
Public Function Decode(ByVal Key As ResourceKey, Bytes() As Byte) As Long
    If Key Is Nothing Then _
        Error.ArgumentNull "Key"
    If CorArray.IsNull(Bytes) Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Key.ResourceType <> ResourceTypes.BitmapResource Then _
        Error.Argument Argument_InvalidResourceKeyType, "Key"
    
    Set mBitmap = CreatePicture(Bytes)
    
    If Not mBitmap Is Nothing Then
        Set mKey = Key
        Decode = 1
    End If
End Function

''
' Returns the of resource this decoder can decode.
'
' @return Returns either a number or string representing the resource type.
' @see ResourceTypes
'
Public Function GetDecodeType() As Variant
    GetDecodeType = ResourceTypes.BitmapResource
End Function

''
' Returns the next decoded resource available in the decoder.
'
' @param ReturnKey This is set to the resource key of the next available resource.
' @param ReturnValue This is set to the resource value of the next available resource.
' @return Returns True if a return resource was available and set, False otherwise.
' @remarks For this resource type only 1 resource is ever returned. Once it is
' returned, the function will return False and the resource cannot be retrieved again.
' @see ResourceKey
'
Public Function GetResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    If mKey Is Nothing Then
        Exit Function
    End If
    
    Set ReturnKey = mKey
    Set ReturnValue = mBitmap
    
    ' Set these to nothing because we only return them once.
    Set mKey = Nothing
    Set mBitmap = Nothing
    
    GetResource = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreatePicture(ByRef Bytes() As Byte) As StdPicture
    Dim ByteIndex       As Long
    Dim Size            As Long
    Dim hMem            As Long
    Dim lpMem           As Long
    Dim Stream          As IUnknown

    ByteIndex = LBound(Bytes)
    Size = Len1D(Bytes)

    If Not HasBitmapHeader(Bytes) Then
        ' if the original data does not have a header
        ' then we will be adding our own, so bump up
        ' the needed memory allocation.
        Size = Size + SIZEOF_BITMAPFILEHEADER
    End If

    ' lets get that memory
    hMem = GlobalAlloc(GMEM_MOVEABLE, Size)

    If hMem Then
        ' we lock the memory and get its location
        ' so we can copy our data to it.
        lpMem = GlobalLock(hMem)

        If lpMem Then
            ' if the original data has a bitmap header
            ' included, then just copy the data to the
            ' allocated memory. Otherwise, we will need
            ' to add our own header first, then add the data.
            If HasBitmapHeader(Bytes) Then
                CopyMemory ByVal lpMem, Bytes(ByteIndex), Size
            Else
                AddBitmapHeader lpMem, Bytes
                CopyMemory ByVal (lpMem + SIZEOF_BITMAPFILEHEADER), Bytes(ByteIndex), Len1D(Bytes)
            End If

            GlobalUnlock lpMem

            ' we'll create a stream using the memory we filled with our
            ' image as a backing store for the stream. Then we'll just
            ' create a picture from the stream.
            If CreateStreamOnHGlobal(hMem, BOOL_FALSE, Stream) = S_OK Then
                Dim UnkIID As VBGUID
                With UnkIID
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With

                Dim Unk As IUnknown
                OleLoadPicture Stream, Size, BOOL_FALSE, UnkIID, Unk
            End If
        End If

        GlobalFree hMem
    End If

    ' Return it, casting to a StdPicture object.
    Set CreatePicture = Unk
End Function

Private Function HasBitmapHeader(ByRef Bytes() As Byte) As Boolean
    Dim ByteIndex As Long
    
    ByteIndex = LBound(Bytes)
    HasBitmapHeader = Bytes(ByteIndex) = vbUpperBChar And Bytes(ByteIndex + 1) = vbUpperMChar
End Function

' The bitmap data begins immediatley after the
' BITMAPFILEHEADER, BITMAPINFOHEADER and any
' RGBQUAD structures. There are no RGBQUAD
' structures for 24bit pictures.
Private Sub AddBitmapHeader(ByVal lpMem As Long, ByRef Bytes() As Byte)
    Dim DataStartOffset As Long
    Dim ByteIndex       As Long
    
    ByteIndex = LBound(Bytes)
    
    ' All offsets include the BITMAPFILEHEADER(14) and BITMAPINFOHEADER(40)
    ' so start the offset with that.
    DataStartOffset = 54
    
    Const BITCOUNT_OFFSET   As Long = 14
    Const COLORSUSED_OFFSET As Long = 32
    
    Dim ColorsUsed As Long
    ColorsUsed = AsLong(Bytes(ByteIndex + COLORSUSED_OFFSET))
    
    ' If ColorsUsed is 0 then the maximum number
    ' of colors is used based on the bits-per-pixel.
    If ColorsUsed = 0 Then
        Dim BitCount As Long
        BitCount = AsWord(Bytes(ByteIndex + BITCOUNT_OFFSET))
        ColorsUsed = Powers(BitCount)
    End If
    
    ' Add the RGBQUAD structure offset.
    DataStartOffset = DataStartOffset + ColorsUsed * 4
    
    ' The BITMAPFILEHEADER gets aligned within
    ' the datatypes which misaligns the datatypes
    ' with what the Stream expects. So, we build
    ' the header by hand.
    Dim BmpHeader(SIZEOF_BITMAPFILEHEADER - 1) As Byte
    
    ' Spell "BM"
    BmpHeader(0) = vbUpperBChar
    BmpHeader(1) = vbUpperMChar
    
    ' Bytes 3-6 contain the size of the file in bytes.
    AsLong(BmpHeader(2)) = CorArray.Length(Bytes) + SIZEOF_BITMAPFILEHEADER
    
    ' Bytes 11-14 contain the offset to the data bytes.
    AsLong(BmpHeader(10)) = DataStartOffset
        
        ' Shove our header onto the Stream before we shove the data.
'        Stream.Write BmpHeader(0), SIZEOF_BITMAPFILEHEADER, 0&
    CopyMemory ByVal lpMem, BmpHeader(0), SIZEOF_BITMAPFILEHEADER
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceDecoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceDecoder_Decode(ByVal Key As ResourceKey, Bytes() As Byte) As Long
    IResourceDecoder_Decode = Decode(Key, Bytes)
End Function

Private Function IResourceDecoder_GetDecodeType() As Variant
    IResourceDecoder_GetDecodeType = GetDecodeType
End Function

Private Function IResourceDecoder_GetResource(ReturnKey As ResourceKey, ReturnValue As Variant) As Boolean
    IResourceDecoder_GetResource = GetResource(ReturnKey, ReturnValue)
End Function
