Attribute VB_Name = "modHelper"
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
' Module: modHelper
'

''
' Creates an object that provides ASM code for special functions.
'
'
Option Explicit

Private Const OutOfMemoryError      As Long = 7
Private Const OffsetToFirstFunction As Long = 4
Private Const Swap4Offset           As Long = 0
Private Const Swap8Offset           As Long = 5
Private Const Swap16Offset          As Long = 13
Private Const Swap2Offset           As Long = 25
Private Const Swap1Offset           As Long = 31
Private Const DerefEBPOffset        As Long = 36
Private Const MoveVariantOffset     As Long = 39
Private Const EcvtCallOffset        As Long = 50
Private Const ShiftRightOffset      As Long = 72
Private Const ShiftLeftOffset       As Long = 76
Private Const UAddOffset            As Long = 80
Private Const UAdd64Offset          As Long = 84


' We expose this directly so we don't have to manage reference counting. This requires
' a call to InitHelper before it can be used. The call is currently made in Main.
Public Helper As IHelper

Private Type HelperVTable
    pVTable     As Long
    Func(17)    As Long
End Type

Private mAsm()      As Long
Private mMSVCLib    As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitHelper()
    Dim Table As HelperVTable
    
    InitAsm
    InitVTable Table
    AllocateObject Table
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitVTable(ByRef Table As HelperVTable)
    With Table
        .Func(0) = FuncAddr(AddressOf QueryInterface)
        .Func(1) = FuncAddr(AddressOf AddRef)
        .Func(2) = FuncAddr(AddressOf Release)
        .Func(3) = VarPtr(mAsm(Swap4Offset))
        .Func(4) = VarPtr(mAsm(Swap8Offset))
        .Func(5) = VarPtr(mAsm(Swap16Offset))
        .Func(6) = VarPtr(mAsm(Swap2Offset))
        .Func(7) = VarPtr(mAsm(Swap1Offset))
        .Func(8) = VarPtr(mAsm(DerefEBPOffset))
        .Func(9) = VarPtr(mAsm(MoveVariantOffset))
        .Func(10) = VarPtr(mAsm(EcvtCallOffset))
        .Func(11) = VarPtr(mAsm(ShiftRightOffset))
        .Func(12) = VarPtr(mAsm(ShiftLeftOffset))
        .Func(13) = VarPtr(mAsm(UAddOffset))
        .Func(14) = VarPtr(mAsm(UAdd64Offset))
    End With
End Sub

Private Sub AllocateObject(ByRef Table As HelperVTable)
    Dim This As Long
    This = CoTaskMemAlloc(LenB(Table))
    If This = vbNullPtr Then _
        Err.Raise OutOfMemoryError
        
    Table.pVTable = This + OffsetToFirstFunction
    CopyMemory ByVal This, Table, LenB(Table)
    
    ObjectPtr(Helper) = This
End Sub

Private Sub InitAsm()
    ReDim mAsm(168)
    ' Swap4  from Matt Curland
    mAsm(0) = &H824448B
    mAsm(1) = &HC24548B
    mAsm(2) = &HA87088B
    mAsm(3) = &HCC20889
    mAsm(4) = &H90909000
        
    ' Swap8
    mAsm(5) = &H824448B
    mAsm(6) = &HC24548B
    mAsm(7) = &HA87088B
    mAsm(8) = &H488B0889
    mAsm(9) = &H44A8704
    mAsm(10) = &HC2044889
    mAsm(11) = &H9090000C
    mAsm(12) = &H90909090
    
    ' Swap16
    mAsm(13) = &H824448B
    mAsm(14) = &HC24548B
    mAsm(15) = &HA87088B
    mAsm(16) = &H488B0889
    mAsm(17) = &H44A8704
    mAsm(18) = &H8B044889
    mAsm(19) = &H4A870848
    mAsm(20) = &H8488908
    mAsm(21) = &H870C488B
    mAsm(22) = &H48890C4A
    mAsm(23) = &HCC20C
    mAsm(24) = &H33909090
        
    ' Swap2
    mAsm(25) = &H824448B
    mAsm(26) = &HC24548B
    mAsm(27) = &H66088B66
    mAsm(28) = &H89660A87
    mAsm(29) = &HCC208
    mAsm(30) = &H33909090

    ' Swap1
    mAsm(31) = &H824448B
    mAsm(32) = &HC24548B
    mAsm(33) = &HA86088A
    mAsm(34) = &HCC20888
    mAsm(35) = &H90909000

    ' DerefEBP  from Matt Curland
    mAsm(36) = &H8244C8B
    mAsm(37) = &HD448B
    mAsm(38) = &H900008C2

    ' MoveVariant from Matt Curland
    mAsm(39) = &HC24448B
    mAsm(40) = &H824548B
    mAsm(41) = &H8B56C88B
    mAsm(42) = &H8B328931
    mAsm(43) = &H72890471
    mAsm(44) = &H8718B04
    mAsm(45) = &H5E087289
    mAsm(46) = &H890C498B
    mAsm(47) = &HC7660C4A
    mAsm(48) = &HC2000000
    mAsm(49) = &H9090000C

    ' _ecvt call
    mAsm(50) = &H81EC8B55
    mAsm(51) = &HC0EC&
    mAsm(52) = &H57565300
    mAsm(53) = &HFF40BD8D
    mAsm(54) = &H30B9FFFF
    mAsm(55) = &HB8000000
    mAsm(56) = &HCCCCCCCC
    mAsm(57) = &H458BABF3
    mAsm(58) = &H4D8B501C
    mAsm(59) = &H558B5118
    mAsm(60) = &H45DD5214
    mAsm(61) = &H8EC830C
    mAsm(62) = &HB8241CDD
    mAsm(63) = &HFFFFF3EC   ' ecvt address goes here
    mAsm(64) = &H9090D0FF
    mAsm(65) = &H5F14C483
    mAsm(66) = &HC4815B5E
    mAsm(67) = &HC0&
    mAsm(68) = &H9090EC3B
    mAsm(69) = &H8B909090
    mAsm(70) = &H18C25DE5
    mAsm(71) = &H90909000
    
    ' compatible libraries
    ' msvcrt20.dll
    ' msvcrt40.dll
    ' msvcr70.dll
    ' msvcr71.dll
    ' msvcr71d.dll
    mMSVCLib = LoadLibraryA("msvcrt.dll")
    mAsm(63) = GetProcAddress(mMSVCLib, "_ecvt")
    
    'shift right
    mAsm(72) = &H824448B
    mAsm(73) = &HC244C8B
    mAsm(74) = &HCC2E8D3
    mAsm(75) = &HCCCCCC00
    
    'shift left
    mAsm(76) = &H824448B
    mAsm(77) = &HC244C8B
    mAsm(78) = &HCC2E0D3
    mAsm(79) = &HCCCCCC00
    
    'UAdd from Matt Curland
    mAsm(80) = &HC24448B
    mAsm(81) = &H8244C8B
    mAsm(82) = &HCC2C103
    mAsm(83) = &H90909000

    'UAdd64
    mAsm(84) = &H1024448B
    mAsm(85) = &H8244C8B
    mAsm(86) = &H548B0101
    mAsm(87) = &H448B1424
    mAsm(88) = &H10010C24
    mAsm(89) = &H3731039
    mAsm(90) = &HC2010183
    mAsm(91) = &HCCCC0014

    Call VirtualProtect(mAsm(0), UBound(mAsm) * 4, PAGE_EXECUTE_READWRITE, 0&)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IUnknown Interface Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function QueryInterface(ByVal This As Long, ByVal riid As Long, pvObj As Long) As Long
    QueryInterface = E_NOINTERFACE
End Function

Private Function AddRef(ByVal This As Long) As Long
    ' do nothing
End Function

Private Function Release(ByVal This As Long) As Long
    CoTaskMemFree This
End Function
