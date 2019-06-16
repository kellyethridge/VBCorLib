VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: EncodingTable
'
Option Explicit

' The binary data read from the BinaryReader is
' in the same order as listed in this type. If more
' encoding items are to be added, they must be added
' in the same order using a BinaryWriter.
Private Type DataItem
    BodyName            As String
    CodePage            As Long
    EncodingName        As String
    HeaderName          As String
    IsBrowserDisplay    As Boolean
    IsBrowserSave       As Boolean
    IsMailNewsDisplay   As Boolean
    IsMailNewsSave      As Boolean
    WebName             As String
    WindowsCodePage     As Long
End Type

Private mEncodingData()     As DataItem
Private mIndexByCodePage    As Hashtable
Private mCodePageByName     As Hashtable


Friend Property Get BodyName(ByVal cp As Long) As String
    BodyName = mEncodingData(GetIndex(cp)).BodyName
End Property

Friend Property Get EncodingName(ByVal cp As Long) As String
    EncodingName = mEncodingData(GetIndex(cp)).EncodingName
End Property

Friend Property Get HeaderName(ByVal cp As Long) As String
    HeaderName = mEncodingData(GetIndex(cp)).HeaderName
End Property

Friend Property Get IsBrowserDisplay(ByVal cp As Long) As Boolean
    IsBrowserDisplay = mEncodingData(GetIndex(cp)).IsBrowserDisplay
End Property

Friend Property Get IsBrowserSave(ByVal cp As Long) As Boolean
    IsBrowserSave = mEncodingData(GetIndex(cp)).IsBrowserSave
End Property

Friend Property Get IsMailNewsDisplay(ByVal cp As Long) As Boolean
    IsMailNewsDisplay = mEncodingData(GetIndex(cp)).IsMailNewsDisplay
End Property

Friend Property Get IsMailNewsSave(ByVal cp As Long) As Boolean
    IsMailNewsSave = mEncodingData(GetIndex(cp)).IsMailNewsSave
End Property

Friend Property Get WebName(ByVal cp As Long) As String
    WebName = mEncodingData(GetIndex(cp)).WebName
End Property

Friend Property Get WindowsCodePage(ByVal cp As Long) As Long
    WindowsCodePage = mEncodingData(GetIndex(cp)).WindowsCodePage
End Property

Friend Function GetCodePage(ByRef Name As String) As Long
    Dim CodePage As Variant
    
    CodePage = mCodePageByName(Name)
    
    If IsEmpty(CodePage) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingNotSupported, Name), "Name")
    
    GetCodePage = CodePage
End Function

Friend Function GetEncodings() As EncodingInfo()
    Dim Encodings() As EncodingInfo
    ReDim Encodings(0 To UBound(mEncodingData))
    
    Dim i As Long
    For i = 0 To UBound(Encodings)
        Set Encodings(i) = New EncodingInfo
        
        With mEncodingData(i)
            Encodings(i).Init .CodePage, .WebName, .EncodingName
        End With
    Next i
    
    GetEncodings = Encodings
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    LoadData
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIndex(ByVal cp As Long) As Long
    If mIndexByCodePage.Contains(cp) Then
        GetIndex = mIndexByCodePage(cp)
    Else
        Throw Cor.NewNotSupportedException("The CodePage is either invalid or not installed.")
    End If
End Function

' The byte stream is formatted using a BinaryWriter.
' The first item written is the number of encodings in
' the stream, as a vbLong. The order to read in each item
' of an encoding is the same as defined in DataItem at the top.
Private Sub LoadData()
    Dim Data()          As Byte
    Dim ms              As MemoryStream
    Dim Reader          As BinaryReader
    Dim NumEncodings    As Long
    
    Data = LoadResData(101, "ENCODINGDATA")
    Set ms = Cor.NewMemoryStream(Data, Writable:=False)
    Set Reader = Cor.NewBinaryReader(ms)
    NumEncodings = Reader.ReadInt32
    
    ReDim mEncodingData(0 To NumEncodings - 1)
    
    Set mIndexByCodePage = New Hashtable
    Set mCodePageByName = Cor.NewHashtable(EqualityComparer:=StringComparer.OrdinalIgnoreCase)
    
    Dim i As Long
    For i = 0 To NumEncodings - 1
        FillData Reader, mEncodingData(i)
        mIndexByCodePage.Add mEncodingData(i).CodePage, i
        mCodePageByName(mEncodingData(i).WebName) = mEncodingData(i).CodePage
        mCodePageByName(mEncodingData(i).BodyName) = mEncodingData(i).CodePage
        mCodePageByName(mEncodingData(i).EncodingName) = mEncodingData(i).CodePage
    Next i
    
    Reader.CloseReader
End Sub

Private Sub FillData(ByVal Reader As BinaryReader, ByRef Data As DataItem)
    With Data
        .BodyName = Reader.ReadString
        .CodePage = Reader.ReadInt32
        .EncodingName = Reader.ReadString
        .HeaderName = Reader.ReadString
        .IsBrowserDisplay = Reader.ReadBoolean
        .IsBrowserSave = Reader.ReadBoolean
        .IsMailNewsDisplay = Reader.ReadBoolean
        .IsMailNewsSave = Reader.ReadBoolean
        .WebName = Reader.ReadString
        .WindowsCodePage = Reader.ReadInt32
    End With
End Sub


