VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICryptoTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: ICryptoTransform
'

''
' A common interface implemented by all crypto transform classes.
'
' @see CryptoAPITransform
' @see RijndaelManagedTransform
' @see HMAC
' @see HMACSHA1
' @see HMACSHA256
' @see HMACSHA384
' @see HMACSHA512
' @see HMACMD5
' @see HMACRIPEMD160
'
Option Explicit

''
' Represents if the transform can be reused to start transorming a new set of data.
'
' @return Returns True if the transform can be reused, False otherwise.
' @remarks When a transform has finished the final transformation, it may support
' starting over to begin tranforming a new set of data.
'
Public Property Get CanReuseTransform() As Boolean: End Property

''
' Returns if a transform can transform multiple blocks in a single call.
'
' @return Returns True if multiple blocks can be transformed, False otherwise.
'
Public Property Get CanTransformMultipleBlocks() As Boolean: End Property

''
' The size of the block of plain text that is to be tranformed.
'
' @return Returns the size of plain text in bytes to be transformed.
'
Public Property Get InputBlockSize() As Long: End Property

''
' The size of the block of ciphered text after tranformation occurred.
'
' @return Returns the size of the ciphered text in bytes.
'
Public Property Get OutputBlockSize() As Long: End Property

''
' Transforms a block of plain text data, placing the ciphered text in the output buffer.
'
' @param InputBuffer The plain text to be transformed.
' @param InputOffset The starting index to begin transforming.
' @param InputCount The number of bytes to be processed.
' @param OutputBuffer The array to place the transformed data.
' @param OutputOffset The starting index to begin placing the transformed data.
' @return The number of byte actually placed in the output buffer during this call.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long: End Function

''
' Transforms the last of the data, adding any padding as needed.
'
' @param InputBuffer The plain text buffer to process.
' @param InputOffset The starting index of the data to be processed.
' @param InputCount The amount of data to be processed.
' @return The final array of processed data.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte(): End Function

