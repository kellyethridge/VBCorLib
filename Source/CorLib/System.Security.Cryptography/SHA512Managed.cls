VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA512Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: SHA512Managed
'

''
' Provides services to compute the SHA-512 hash value.
'
' @remarks The SHA-512 algorithm is implemented in Visual Basic.
' <p>Details can be found at http://csrc.nist.gov/CryptoToolkit/tkhash.html</p>
' @see HashAlgorithm
'
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements ICryptoTransform
Implements SHA512
Implements IHashAlgorithm

Private mBase               As HashAlgorithmBase
Private mWh(0 To 79)        As Long
Private mWl(0 To 79)        As Long
Private mStateH(0 To 7)     As Long
Private mStateL(0 To 7)     As Long
Private mKh(0 To 79)        As Long
Private mKl(0 To 79)        As Long
Private mStatePtr           As DLong


''
' Returns if this service provider can be reused to transfrom more data.
'
' @return Always returns True.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if this transform can operate on multiple blocks at once.
'
' @return Always returns True.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = True
End Property

''
' Returns the computed hash.
'
' @return A 64-byte array containing the hash value.
' @remarks The TransformFinalBlock must be called before accessing
' this property or an InvalidOperationException will be thrown.
'
Public Property Get Hash() As Byte()
    Hash = mBase.Hash
End Property

''
' Returns the size of the computed hash.
'
' @return Always returns 512 bits.
'
Public Property Get HashSize() As Long
    HashSize = 512
End Property

''
' Returns the input block size this transform can work on.
'
' @return Always returns 1.
' @remarks When calling the TransformBlock method, the InputCount must be zero or a multiple of this value.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = 1
End Property

''
' Returns the output block size this transform can produce.
'
' @return Always returns 1.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = 1
End Property

''
' Releases any resources.
'
Public Sub Clear()
    mBase.Clear
End Sub

''
' Computes a hash value for a source in a single pass.
'
' @param Source This can be either a Byte array or any Stream compatible object.
' @param Index The starting index in the Byte array. This is ignored for Stream objects.
' @param Count The number of bytes in the array to be computed. This is ignored for Stream objects.
' @return A 64-byte hash value.
' @remarks If a Stream object is passed in, then the stream is continuously read and the hash calculated until
' there is no more data left to read from the stream.
'
Public Function ComputeHash(ByRef Source As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    ComputeHash = mBase.ComputeHash(Source, Index, Count)
End Function

''
' Resets the hashing algorithm to start again.
'
Public Sub Initialize()
    Erase mWh
    Erase mWl
    
    CopyMemory mStateH(0), ByVal mStatePtr.HiDWord, 32
    CopyMemory mStateL(0), ByVal mStatePtr.LoDWord, 32
    
    mBase.Initialize
End Sub

''
' Computes the hash for the specified block of data.
'
' @param InputBuffer The data to compute the hash from.
' @param InputOffset The starting index in the input data to compute the hash.
' @param InputCount The number of bytes to compute the hash from.
' @param OutputBuffer The array to place the input buffer bytes in.
' @param OutputOffset The starting index to beging copying the bytes to.
' @return The number of bytes processed.
' @remarks The OutputBuffer will contain the same data as the InputBuffer. No hash values are placed in the OutputBuffer.
' <p>If the OutputBuffer is a Null Array, or is the same array as the InputBuffer and the OutputOffset equals the InputOffset, then nothing is copied, however, the hash is still computed
' on the InputBuffer data.</p>
' <p>The TransformFinalBlock needs to be called to finish computing the hash, this method alone cannot compute the entire hash.</p>
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    TransformBlock = mBase.TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

''
' Computes the final hash value.
'
' @param InputBuffer The final block of data to compute the hash for.
' @param InputOffset The index in the InputBuffer to begin computing data.
' @param InputCount The number of bytes to be processed.
' @return Returns the section of the InputBuffer that was processed.
' @remarks This does not return the computed hash value. A copy of the data that was
' process is returned. To retrieve the hash value, call the Hash property.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    TransformFinalBlock = mBase.TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare against this instance.
' @return Returns True if the values are the same.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Returns a number identifing this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return Returns a string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    With Statics.SHA512
        CopyMemory mKh(0), ByVal .KPtr.HiDWord, 320
        CopyMemory mKl(0), ByVal .KPtr.LoDWord, 320
        mStatePtr = .StatePtr
    End With
    
    Set mBase = Cor.NewHashAlgorithmBase(Me, True, True, 128)
    Initialize
End Sub

Friend Sub Init(ByRef StatePtr As DLong)
    mStatePtr = StatePtr
    Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashFinal(ByRef Buffer() As Byte, ByVal BufferCount As Long, ByVal MessageLength As Currency) As Byte()
    Buffer(BufferCount) = &H80
    BufferCount = BufferCount + 1

    Dim Size As Long
    Size = IIf(BufferCount > 112, 256, 128)

    Dim FinalBuffer() As Byte
    ReDim FinalBuffer(0 To Size - 1)
    CopyMemory FinalBuffer(0), Buffer(0), BufferCount

    Dim i As Long
    With AsOByteArr(MessageLength * 0.0008@)
        For i = 0 To 7
            FinalBuffer(Size - i - 1) = .Bytes(i)
        Next i
    End With

    Transform FinalBuffer, 0
    If Size = 256 Then
        Transform FinalBuffer, 128
    End If

    Dim Result() As Byte
    ReDim Result(0 To 63)
    For i = 0 To 7
        AsLong(Result(i * 8)) = SwapEndian(mStateH(i))
        AsLong(Result(i * 8 + 4)) = SwapEndian(mStateL(i))
    Next i

    HashFinal = Result
End Function

#If Release Then
' Must disable Integer overflow checks.
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim i As Long
    For i = 0 To 15
        mWh(i) = (Bytes(Index) * &H1000000) Or (Bytes(Index + 1) * &H10000) Or (Bytes(Index + 2) * &H100&) Or Bytes(Index + 3)
        Index = Index + 4
        mWl(i) = (Bytes(Index) * &H1000000) Or (Bytes(Index + 1) * &H10000) Or (Bytes(Index + 2) * &H100&) Or Bytes(Index + 3)
        Index = Index + 4
    Next i
    
    Dim Lo As Long
    Dim th As Long: Dim tl As Long
    Dim sh As Long: Dim sl As Long
    
    For i = 16 To 79
        '
        ' sum = LittleSigma1(W(i-2)) + W(i-7)
        '
        tl = mWl(i - 2)
        th = mWh(i - 2)
        sh = (((((th And &HFFF80000) \ &H80000) And &H1FFF&) Or (tl * &H2000&)) Xor _
             ((th * &H8&) Or (((tl And &HE0000000) \ &H20000000) And &H7&)) Xor _
             (((th And &HFFFFFFC0) \ &H40&) And &H3FFFFFF)) + _
             mWh(i - 7)
        sl = (((((tl And &HFFF80000) \ &H80000) And &H1FFF&) Or (th * &H2000&)) Xor _
             ((tl * &H8&) Or (((th And &HE0000000) \ &H20000000) And &H7&)) Xor _
             ((th * &H4000000) Or (((tl And &HFFFFFFC0) \ &H40&) And &H3FFFFFF))) + _
             mWl(i - 7)
        If (sl - &H80000000) < (mWl(i - 7) - &H80000000) Then sh = sh + 1

        '
        ' sum = sum + LittleSigma0(W(i-15))
        '
        tl = mWl(i - 15)
        th = mWh(i - 15)
        sh = sh + (((tl * &H80000000) Or (((th And &HFFFFFFFE) \ &H2) And &H7FFFFFFF)) Xor _
                  ((tl * &H1000000) Or (((th And &HFFFFFF00) \ &H100&) And &HFFFFFF)) Xor _
                  (((th And &HFFFFFF80) \ &H80&) And &H1FFFFFF))
        Lo = ((th * &H80000000) Or (((tl And &HFFFFFFFE) \ &H2) And &H7FFFFFFF)) Xor _
             ((th * &H1000000) Or (((tl And &HFFFFFF00) \ &H100&) And &HFFFFFF)) Xor _
             ((th * &H2000000) Or (((tl And &HFFFFFF80) \ &H80&) And &H1FFFFFF))
        sl = sl + Lo
        If (sl - &H80000000) < (Lo - &H80000000) Then sh = sh + 1
        
        '
        ' sum = sum + W(i-16)
        '
        sh = sh + mWh(i - 16)
        sl = sl + mWl(i - 16)
        If (sl - &H80000000) < (mWl(i - 16) - &H80000000) Then sh = sh + 1
        
        '
        ' W(i) = sum
        '
        mWh(i) = sh
        mWl(i) = sl
    Next i
    
    Dim Ah As Long, Al As Long
    Dim Bh As Long, Bl As Long
    Dim Ch As Long, Cl As Long
    Dim Dh As Long, Dl As Long
    Dim Eh As Long, El As Long
    Dim Fh As Long, Fl As Long
    Dim Gh As Long, Gl As Long
    Dim Hh As Long, Hl As Long
    
    Ah = mStateH(0): Al = mStateL(0)
    Bh = mStateH(1): Bl = mStateL(1)
    Ch = mStateH(2): Cl = mStateL(2)
    Dh = mStateH(3): Dl = mStateL(3)
    Eh = mStateH(4): El = mStateL(4)
    Fh = mStateH(5): Fl = mStateL(5)
    Gh = mStateH(6): Gl = mStateL(6)
    Hh = mStateH(7): Hl = mStateL(7)
    
    
    Dim T1h As Long: Dim T1l As Long
    Dim T2h As Long: Dim T2l As Long
    
    For i = 0 To 79
        
        'BigSigma0  : ROTR(x,28) XOR ROTR(x,34) XOR ROTR(x,39)
        'BigSigma1  : ROTR(x,14) XOR ROTR(x,18) XOR ROTR(x,41)
        'Ch         : (x AND y) XOR ((NOT x) AND z)
        'Maj        : (x AND y) XOR (x AND z) XOR (y AND z)
        '
        'T1 = H + BigSigma1(E) + Ch(E,F,G) + K(i) + W(i)
        'T2 = BigSigma0(A) + Maj(A,B,C)
        '
        
        '
        ' T1 = H + BigSigma1(E)
        '
        T1h = Hh + _
              (((((Eh And &HFFFFC000) \ &H4000) And &H3FFFF) Or (El * &H40000)) Xor _
              ((((Eh And &HFFFC0000) \ &H40000) And &H3FFF) Or (El * &H4000)) Xor _
              ((Eh * &H800000) Or (((El And &HFFFFFE00) \ &H200&) And &H7FFFFF)))
        Lo = (((((El And &HFFFFC000) \ &H4000) And &H3FFFF) Or (Eh * &H40000)) Xor _
             ((((El And &HFFFC0000) \ &H40000) And &H3FFF) Or (Eh * &H4000)) Xor _
             ((El * &H800000) Or (((Eh And &HFFFFFE00) \ &H200&) And &H7FFFFF)))
        T1l = Hl + Lo
        If (T1l - &H80000000) < (Lo - &H80000000) Then T1h = T1h + 1
        
        '
        ' T1 = T1 + Ch(E,F,G)
        '
        T1h = T1h + ((Eh And Fh) Xor ((Not Eh) And Gh))
        Lo = ((El And Fl) Xor ((Not El) And Gl))
        T1l = T1l + Lo
        If (T1l - &H80000000) < (Lo - &H80000000) Then T1h = T1h + 1
        
        '
        ' T1 = T1 + K(i)
        '
        T1h = T1h + mKh(i)
        T1l = T1l + mKl(i)
        If (T1l - &H80000000) < (mKl(i) - &H80000000) Then T1h = T1h + 1
        
        '
        ' T1 = T1 + W(i)
        '
        T1h = T1h + mWh(i)
        T1l = T1l + mWl(i)
        If (T1l - &H80000000) < (mWl(i) - &H80000000) Then T1h = T1h + 1
        
        '
        ' T2 = BigSigma0(A)
        '
        T2h = ((((Ah And &HF0000000) \ &H10000000) And &HF&) Or (Al * &H10&)) Xor _
              ((Ah * &H40000000) Or (((Al And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)) Xor _
              ((Ah * &H2000000) Or (((Al And &HFFFFFF80) \ &H80&) And &H1FFFFFF))
        T2l = ((((Al And &HF0000000) \ &H10000000) And &HF&) Or (Ah * &H10&)) Xor _
              ((Al * &H40000000) Or (((Ah And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)) Xor _
              ((Al * &H2000000) Or (((Ah And &HFFFFFF80) \ &H80&) And &H1FFFFFF))
        
        '
        ' T2 = T2 + Maj(A,B,C)
        '
        T2h = T2h + ((Ah And Bh) Xor (Ah And Ch) Xor (Bh And Ch))
        Lo = (Al And Bl) Xor (Al And Cl) Xor (Bl And Cl)
        T2l = T2l + Lo
        If (T2l - &H80000000) < (Lo - &H80000000) Then T2h = T2h + 1
        
        Hh = Gh: Hl = Gl
        Gh = Fh: Gl = Fl
        Fh = Eh: Fl = El
        Eh = Dh + T1h
        El = Dl + T1l
        If (El - &H80000000) < (T1l - &H80000000) Then Eh = Eh + 1
        
        Dh = Ch: Dl = Cl
        Ch = Bh: Cl = Bl
        Bh = Ah: Bl = Al
        Ah = T1h + T2h
        Al = T1l + T2l
        If (Al - &H80000000) < (T2l - &H80000000) Then Ah = Ah + 1
    Next i
    
    mStateH(0) = mStateH(0) + Ah: mStateL(0) = mStateL(0) + Al: If (mStateL(0) - &H80000000) < (Al - &H80000000) Then mStateH(0) = mStateH(0) + 1
    mStateH(1) = mStateH(1) + Bh: mStateL(1) = mStateL(1) + Bl: If (mStateL(1) - &H80000000) < (Bl - &H80000000) Then mStateH(1) = mStateH(1) + 1
    mStateH(2) = mStateH(2) + Ch: mStateL(2) = mStateL(2) + Cl: If (mStateL(2) - &H80000000) < (Cl - &H80000000) Then mStateH(2) = mStateH(2) + 1
    mStateH(3) = mStateH(3) + Dh: mStateL(3) = mStateL(3) + Dl: If (mStateL(3) - &H80000000) < (Dl - &H80000000) Then mStateH(3) = mStateH(3) + 1
    mStateH(4) = mStateH(4) + Eh: mStateL(4) = mStateL(4) + El: If (mStateL(4) - &H80000000) < (El - &H80000000) Then mStateH(4) = mStateH(4) + 1
    mStateH(5) = mStateH(5) + Fh: mStateL(5) = mStateL(5) + Fl: If (mStateL(5) - &H80000000) < (Fl - &H80000000) Then mStateH(5) = mStateH(5) + 1
    mStateH(6) = mStateH(6) + Gh: mStateL(6) = mStateL(6) + Gl: If (mStateL(6) - &H80000000) < (Gl - &H80000000) Then mStateH(6) = mStateH(6) + 1
    mStateH(7) = mStateH(7) + Hh: mStateL(7) = mStateL(7) + Hl: If (mStateL(7) - &H80000000) < (Hl - &H80000000) Then mStateH(7) = mStateH(7) + 1
End Sub
#Else
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim i As Long
    For i = 0 To 15
        mWh(i) = SwapEndian(AsLong(Bytes(Index)))
        mWl(i) = SwapEndian(AsLong(Bytes(Index + 4)))
        Index = Index + 8
    Next i
    
    Dim T1l As Long
    Dim T1h As Long
    Dim T2l As Long
    Dim T2h As Long
    Dim s0l As Long
    Dim s0h As Long
    Dim s1l As Long
    Dim s1h As Long
    For i = 16 To 79
        T1l = mWl(i - 2)
        T1h = mWh(i - 2)
        s1l = LilSig1Lo(T1h, T1l)
        s1h = LilSig1Hi(T1h, T1l)

        T2l = mWl(i - 15)
        T2h = mWh(i - 15)
        s0l = LilSig0Lo(T2h, T2l)
        s0h = LilSig0Hi(T2h, T2l)
        
        UAdd64 s1h, s1l, mWh(i - 7), mWl(i - 7)
        UAdd64 s1h, s1l, s0h, s0l
        UAdd64 s1h, s1l, mWh(i - 16), mWl(i - 16)

        mWh(i) = s1h
        mWl(i) = s1l
    Next i
    
    Dim Ah As Long, Al As Long
    Dim Bh As Long, Bl As Long
    Dim Ch As Long, Cl As Long
    Dim Dh As Long, Dl As Long
    Dim Eh As Long, El As Long
    Dim Fh As Long, Fl As Long
    Dim Gh As Long, Gl As Long
    Dim Hh As Long, Hl As Long
    
    Ah = mStateH(0): Al = mStateL(0)
    Bh = mStateH(1): Bl = mStateL(1)
    Ch = mStateH(2): Cl = mStateL(2)
    Dh = mStateH(3): Dl = mStateL(3)
    Eh = mStateH(4): El = mStateL(4)
    Fh = mStateH(5): Fl = mStateL(5)
    Gh = mStateH(6): Gl = mStateL(6)
    Hh = mStateH(7): Hl = mStateL(7)
    
    For i = 0 To 79
        T1h = Hh: T1l = Hl
        UAdd64 T1h, T1l, BigSIG1Hi(Eh, El), BigSIG1Lo(Eh, El)
        UAdd64 T1h, T1l, Choice(Eh, Fh, Gh), Choice(El, Fl, Gl)
        UAdd64 T1h, T1l, mKh(i), mKl(i)
        UAdd64 T1h, T1l, mWh(i), mWl(i)

        T2h = BigSIG0Hi(Ah, Al): T2l = BigSIG0Lo(Ah, Al)
        UAdd64 T2h, T2l, Maj(Ah, Bh, Ch), Maj(Al, Bl, Cl)

        Hh = Gh: Hl = Gl
        Gh = Fh: Gl = Fl
        Fh = Eh: Fl = El
        Eh = Dh: El = Dl
        UAdd64 Eh, El, T1h, T1l
        Dh = Ch: Dl = Cl
        Ch = Bh: Cl = Bl
        Bh = Ah: Bl = Al
        Ah = T1h: Al = T1l
        UAdd64 Ah, Al, T2h, T2l
    Next i
    
    UAdd64 mStateH(0), mStateL(0), Ah, Al
    UAdd64 mStateH(1), mStateL(1), Bh, Bl
    UAdd64 mStateH(2), mStateL(2), Ch, Cl
    UAdd64 mStateH(3), mStateL(3), Dh, Dl
    UAdd64 mStateH(4), mStateL(4), Eh, El
    UAdd64 mStateH(5), mStateL(5), Fh, Fl
    UAdd64 mStateH(6), mStateL(6), Gh, Gl
    UAdd64 mStateH(7), mStateL(7), Hh, Hl
End Sub

Private Sub UAdd64(ByRef xh As Long, ByRef xl As Long, ByVal yh As Long, ByVal yl As Long)
    Helper.UAdd64 xh, xl, yh, yl
End Sub

Private Function RShiftHi(ByVal xh As Long, ByVal xl As Long, ByVal Count As Long) As Long
    If Count < 32 Then
        RShiftHi = Helper.ShiftRight(xh, Count)
    End If
End Function

Private Function RShiftLo(ByVal xh As Long, ByVal xl As Long, ByVal Count As Long) As Long
    If Count < 32 Then
        RShiftLo = Helper.ShiftLeft(xh, 32 - Count) Or Helper.ShiftRight(xl, Count)
    Else
        RShiftLo = Helper.ShiftRight(xh, Count - 32)
    End If
End Function

Private Function RRotateHi(ByVal xh As Long, ByVal xl As Long, ByVal Count As Long) As Long
    If Count < 32 Then
        RRotateHi = Helper.ShiftRight(xh, Count) Or Helper.ShiftLeft(xl, 32 - Count)
    Else
        RRotateHi = Helper.ShiftLeft(xh, 64 - Count) Or Helper.ShiftRight(xl, Count - 32)
    End If
End Function

Private Function RRotateLo(ByVal xh As Long, ByVal xl As Long, ByVal Count As Long) As Long
    If Count < 32 Then
        RRotateLo = Helper.ShiftLeft(xh, 32 - Count) Or Helper.ShiftRight(xl, Count)
    Else
        RRotateLo = Helper.ShiftRight(xh, Count - 32) Or Helper.ShiftLeft(xl, 64 - Count)
    End If
End Function

Private Function Choice(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    Choice = (X And Y) Xor ((Not X) And z)
End Function

Private Function Maj(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    Maj = (X And Y) Xor (X And z) Xor (Y And z)
End Function

Private Function BigSIG0Hi(ByVal xh As Long, ByVal xl As Long) As Long
    BigSIG0Hi = RRotateHi(xh, xl, 28) Xor RRotateHi(xh, xl, 34) Xor RRotateHi(xh, xl, 39)
End Function

Private Function BigSIG0Lo(ByVal xh As Long, ByVal xl As Long) As Long
    BigSIG0Lo = RRotateLo(xh, xl, 28) Xor RRotateLo(xh, xl, 34) Xor RRotateLo(xh, xl, 39)
End Function

Private Function BigSIG1Hi(ByVal xh As Long, ByVal xl As Long) As Long
    BigSIG1Hi = RRotateHi(xh, xl, 14) Xor RRotateHi(xh, xl, 18) Xor RRotateHi(xh, xl, 41)
End Function

Private Function BigSIG1Lo(ByVal xh As Long, ByVal xl As Long) As Long
    BigSIG1Lo = RRotateLo(xh, xl, 14) Xor RRotateLo(xh, xl, 18) Xor RRotateLo(xh, xl, 41)
End Function

Private Function LilSig0Hi(ByVal xh As Long, ByVal xl As Long) As Long
    LilSig0Hi = RRotateHi(xh, xl, 1) Xor RRotateHi(xh, xl, 8) Xor RShiftHi(xh, xl, 7)
End Function

Private Function LilSig0Lo(ByVal xh As Long, ByVal xl As Long) As Long
    LilSig0Lo = RRotateLo(xh, xl, 1) Xor RRotateLo(xh, xl, 8) Xor RShiftLo(xh, xl, 7)
End Function

Private Function LilSig1Hi(ByVal xh As Long, ByVal xl As Long) As Long
    LilSig1Hi = RRotateHi(xh, xl, 19) Xor RRotateHi(xh, xl, 61) Xor RShiftHi(xh, xl, 6)
End Function

Private Function LilSig1Lo(ByVal xh As Long, ByVal xl As Long) As Long
    LilSig1Lo = RRotateLo(xh, xl, 19) Xor RRotateLo(xh, xl, 61) Xor RShiftLo(xh, xl, 6)
End Function
#End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   HashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HashAlgorithm_CanReuseTransform() As Boolean
    HashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get HashAlgorithm_CanTransformMultipleBlocks() As Boolean
    HashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub HashAlgorithm_Clear()
    Clear
End Sub

Private Function HashAlgorithm_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    HashAlgorithm_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function HashAlgorithm_Equals(Value As Variant) As Boolean
    HashAlgorithm_Equals = Equals(Value)
End Function

Private Function HashAlgorithm_GetHashCode() As Long
    HashAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get HashAlgorithm_Hash() As Byte()
    HashAlgorithm_Hash = Hash
End Property

Private Property Get HashAlgorithm_HashSize() As Long
    HashAlgorithm_HashSize = HashSize
End Property

Private Sub HashAlgorithm_Initialize()
    Initialize
End Sub

Private Property Get HashAlgorithm_InputBlockSize() As Long
    HashAlgorithm_InputBlockSize = InputBlockSize
End Property

Private Property Get HashAlgorithm_OutputBlockSize() As Long
    HashAlgorithm_OutputBlockSize = OutputBlockSize
End Property

Private Function HashAlgorithm_ToString() As String
    HashAlgorithm_ToString = ToString
End Function

Private Function HashAlgorithm_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    HashAlgorithm_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function HashAlgorithm_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    HashAlgorithm_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ICryptoTransform_CanReuseTransform() As Boolean
    ICryptoTransform_CanReuseTransform = CanReuseTransform
End Property

Private Property Get ICryptoTransform_CanTransformMultipleBlocks() As Boolean
    ICryptoTransform_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Property Get ICryptoTransform_InputBlockSize() As Long
    ICryptoTransform_InputBlockSize = InputBlockSize
End Property

Private Property Get ICryptoTransform_OutputBlockSize() As Long
    ICryptoTransform_OutputBlockSize = OutputBlockSize
End Property

Private Function ICryptoTransform_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ICryptoTransform_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function ICryptoTransform_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    ICryptoTransform_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IHashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IHashAlgorithm_HashCore(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Transform Bytes, Index
End Sub

Private Function IHashAlgorithm_HashFinal(ByRef Buffer() As Byte, ByVal BufferCount As Long, ByVal MessageLength As Currency) As Byte()
    IHashAlgorithm_HashFinal = HashFinal(Buffer, BufferCount, MessageLength)
End Function

Private Sub IHashAlgorithm_Initialize()
    Initialize
End Sub


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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SHA512
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SHA512_CanReuseTransform() As Boolean
    SHA512_CanReuseTransform = CanReuseTransform
End Property

Private Property Get SHA512_CanTransformMultipleBlocks() As Boolean
    SHA512_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub SHA512_Clear()
    Clear
End Sub

Private Function SHA512_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    SHA512_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function SHA512_Equals(Value As Variant) As Boolean
    SHA512_Equals = Equals(Value)
End Function

Private Function SHA512_GetHashCode() As Long
    SHA512_GetHashCode = GetHashCode
End Function

Private Property Get SHA512_Hash() As Byte()
    SHA512_Hash = Hash
End Property

Private Property Get SHA512_HashSize() As Long
    SHA512_HashSize = HashSize
End Property

Private Sub SHA512_Initialize()
    Initialize
End Sub

Private Property Get SHA512_InputBlockSize() As Long
    SHA512_InputBlockSize = InputBlockSize
End Property

Private Property Get SHA512_OutputBlockSize() As Long
    SHA512_OutputBlockSize = OutputBlockSize
End Property

Private Function SHA512_ToString() As String
    SHA512_ToString = ToString
End Function

Private Function SHA512_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    SHA512_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function SHA512_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    SHA512_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function




