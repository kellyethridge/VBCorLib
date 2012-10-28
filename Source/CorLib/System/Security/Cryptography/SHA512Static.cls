VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA512Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: SHA512Static
'

''
' Provides methods associated with the SHA-512 class.
'
Option Explicit

Private mKh()       As Long
Private mKl()       As Long
Private mKPtr       As DLong
Private mStateH()   As Long
Private mStateL()   As Long
Private mStatePtr   As DLong



''
' Creates a new instance of an SHA-512 implementation.
'
' @param AlgName The name of the implementation to create an instance of.
' @return A new SHA-512 compatible implementation.
' @remarks If not name is given, then a new SHA512Managed is returned.
'
Public Function Create(Optional ByVal AlgName As String = "SHA512") As SHA512
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get KPtr() As DLong
    If mKPtr.LoDWord = vbNullPtr Then
        mKl = Cor.NewLongs(&HD728AE22, &H23EF65CD, &HEC4D3B2F, &H8189DBBC, &HF348B538, &HB605D019, &HAF194F9B, &HDA6D8118, _
                           &HA3030242, &H45706FBE, &H4EE4B28C, &HD5FFB4E2, &HF27B896F, &H3B1696B1, &H25C71235, &HCF692694, _
                           &H9EF14AD2, &H384F25E3, &H8B8CD5B5, &H77AC9C65, &H592B0275, &H6EA6E483, &HBD41FBD4, &H831153B5, _
                           &HEE66DFAB, &H2DB43210, &H98FB213F, &HBEEF0EE4, &H3DA88FC2, &H930AA725, &HE003826F, &HA0E6E70, _
                           &H46D22FFC, &H5C26C926, &H5AC42AED, &H9D95B3DF, &H8BAF63DE, &H3C77B2A8, &H47EDAEE6, &H1482353B, _
                           &H4CF10364, &HBC423001, &HD0F89791, &H654BE30, &HD6EF5218, &H5565A910, &H5771202A, &H32BBD1B8, _
                           &HB8D2D0C8, &H5141AB53, &HDF8EEB99, &HE19B48A8, &HC5C95A63, &HE3418ACB, &H7763E373, &HD6B2B8A3, _
                           &H5DEFB2FC, &H43172F60, &HA1F0AB72, &H1A6439EC, &H23631E28, &HDE82BDE9, &HB2C67915, &HE372532B, _
                           &HEA26619C, &H21C0C207, &HCDE0EB1E, &HEE6ED178, &H72176FBA, &HA2C898A6, &HBEF90DAE, &H131C471B, _
                           &H23047D84, &H40C72493, &H15C9BEBC, &H9C100D4C, &HCB3E42B6, &HFC657E2A, &H3AD6FAEC, &H4A475817)
        
        mKh = Cor.NewLongs(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
                           &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
                           &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
                           &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
                           &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
                           &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
                           &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
                           &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2, _
                           &HCA273ECE, &HD186B8C7, &HEADA7DD6, &HF57D4F7F, &H6F067AA, &HA637DC5, &H113F9804, &H1B710B35, _
                           &H28DB77F5, &H32CAAB7B, &H3C9EBE0A, &H431D67C4, &H4CC5D4BE, &H597F299C, &H5FCB6FAB, &H6C44198C)
        
        mKPtr.HiDWord = VarPtr(mKh(0))
        mKPtr.LoDWord = VarPtr(mKl(0))
    End If
    
    KPtr = mKPtr
End Property

Friend Property Get StatePtr() As DLong
    If mStatePtr.LoDWord = vbNullPtr Then
        mStateL = Cor.NewLongs(&HF3BCC908, &H84CAA73B, &HFE94F82B, &H5F1D36F1, &HADE682D1, &H2B3E6C1F, &HFB41BD6B, &H137E2179)
        mStateH = Cor.NewLongs(&H6A09E667, &HBB67AE85, &H3C6EF372, &HA54FF53A, &H510E527F, &H9B05688C, &H1F83D9AB, &H5BE0CD19)
        
        mStatePtr.HiDWord = VarPtr(mStateH(0))
        mStatePtr.LoDWord = VarPtr(mStateL(0))
    End If
    
    StatePtr = mStatePtr
End Property

