VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityElementStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: SecurityElementStatic
'

''
' Provides static functions associated to the SecurityElement class.
'
' @see SecurityElement
'
Option Explicit

Private Const vbLeftCheveron    As Long = 60
Private Const vbRightCheveron   As Long = 62

Private mTextFormatter      As New StringBuilder
Private mXml                As New StringBuilder
Private mIndex              As Long



''
' Replaces the set of invalid XML characters in a string to their xml representations.
'
' @param Text The text containing the invalid XML characters.
' @return The text with the valid XML representation.
' @remarks This replaces the following characters:
' <pre>
' "&" = &amp;
' "<" = &lt;
' ">" = &gt;
' Double Quote = &quot;
' Single Quote = &apos;
' </pre>
'
Public Function Escape(ByRef Text As String) As String
    Dim sb As StringBuilder
    Set sb = StringBuilderCache.Acquire _
                                .Append(Text) _
                                .Replace("&", "&amp;") _
                                .Replace("<", "&lt;") _
                                .Replace(">", "&gt;") _
                                .Replace("""", "&quot;") _
                                .Replace("'", "&apos;")

    Escape = StringBuilderCache.GetStringAndRelease(sb)
End Function

''
' Determines if a tag name is valid.
'
' @param Tag The tag to be tested.
' @return Returns True if the tag is valid, False otherwise.
' @remarks A tag is not valid if it contains any of the following: "<", ">", or a space.
'
Public Function IsValidTag(ByRef Tag As String) As Boolean
    IsValidTag = CorString.AllChars(Tag, AddressOf IsValidTagChar)
End Function

''
' Determines if an XML text is valid.
'
' @param Text The text to be tested.
' @return Returns True if the text is valid, False otherwise.
' @remarks The text is not valid if it contains any of the following: "<", ">".
'
Public Function IsValidText(ByRef Text As String) As Boolean
    IsValidText = CorString.AllChars(Text, AddressOf IsValidTextChar)
End Function

''
' Determines if an attribute name if valid.
'
' @param Name The name of the attribute to test.
' @return Returns True if the attribute name is valid, False otherwise.
' @remarks The name if not valid if it contains any of the following: "<", ">", or a space.
'
Public Function IsValidAttributeName(ByRef Name As String) As Boolean
    IsValidAttributeName = IsValidTag(Name)
End Function

''
' Determines if an attribute value is valid.
'
' @param Value The attribute value to test.
' @return Returns True if the value is valid, False otherwise.
' @remarks The value is not valid if it contains any of the following: "<", ">".
'
Public Function IsValidAttributeValue(ByRef Value As String) As Boolean
    IsValidAttributeValue = CorString.AllChars(Value, AddressOf IsValidAttributeValueChar)
End Function

''
' Creates a SecurityElement from an XML string.
'
' @param Xml The string to parse into a SecurityElement.
' @return A SecurityElement representation of the XML string.
'
Public Function FromString(ByVal Xml As String) As SecurityElement
    mXml.Length = 0
    Call mXml.AppendString(Xml)
    mIndex = 0
    
    Set FromString = StartElement
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StartElement() As SecurityElement
    Call ConsumeWhiteSpace

    ' If we reached the end at this point
    ' then we are ok since we didn't actually
    ' start this element yet.
    If mIndex = mXml.Length Then Exit Function
    
    ' We expected to start an element tag of some sort.
    If mXml(mIndex) <> vbLeftCheveron Then Throw New XmlSyntaxException
    
    ' We store this index incase we need to back out.
    ' We don't want to be passed start of a tag if we back out.
    Dim OldIndex As Long
    OldIndex = mIndex
    
    ' Skip the cheveron
    mIndex = mIndex + 1
    
    ' Get to the tag name.
    Call ConsumeWhiteSpace
    
    ' we may be inside a closing tag,
    ' so clean up and get out.
    If mXml(mIndex) = corForwardSlashChar Then
        mIndex = OldIndex
        Exit Function
    End If
    
    Dim Element As New SecurityElement
    With Element
        ' We always have one of these.
        .Tag = GetName
        
        ' Get to the next set of characters.
        Call ConsumeWhiteSpace
        
        ' Lets get attributes if they exist.
        Set .Attributes = GetAttributes
        
        ' Get to the end of the tag.
        Call ConsumeWhiteSpace
        
        Dim EarlyEnd As Boolean
        If mXml(mIndex) = corForwardSlashChar Then
            EarlyEnd = True
            mIndex = mIndex + 1
        End If
        
        Call ConsumeWhiteSpace
                
        ' Must be a '>'
        If mXml(mIndex) <> vbRightCheveron Then Throw New XmlSyntaxException
                
        ' Move passed the '>'
        mIndex = mIndex + 1
                
        If Not EarlyEnd Then
            .Text = GetText
            Set .Children = GetChildren
            Call ConsumeWhiteSpace
            Call EndElement(.Tag)
        End If
    End With

    Set StartElement = Element
End Function

Private Sub EndElement(ByVal Name As String)
    If mIndex = mXml.Length Then Throw New XmlSyntaxException
    
    ' must be a '<'
    If mXml(mIndex) <> vbLeftCheveron Then Throw New XmlSyntaxException
    mIndex = mIndex + 1
    
    Call ConsumeWhiteSpace
    
    ' must be a '/'
    If mXml(mIndex) <> corForwardSlashChar Then Throw New XmlSyntaxException
    mIndex = mIndex + 1
    
    Call ConsumeWhiteSpace
    
    Dim StartIndex As Long
    StartIndex = mIndex
    
    Do
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        
        Select Case mXml(mIndex)
            Case corSpaceChar, vbRightCheveron: Exit Do
        End Select
        
        mIndex = mIndex + 1
    Loop
    
    If StrComp(Name, mXml.ToString(StartIndex, mIndex - StartIndex), vbTextCompare) <> 0 Then Throw New XmlSyntaxException
    
    Call ConsumeWhiteSpace
    
    If mIndex = mXml.Length Then Throw New XmlSyntaxException
    If mXml(mIndex) <> vbRightCheveron Then Throw New XmlSyntaxException
    mIndex = mIndex + 1
End Sub

Private Function GetChildren() As ArrayList
    Dim Ret As New ArrayList
    
    Do
        Dim Child As SecurityElement
        Set Child = StartElement
        If Child Is Nothing Then Exit Do
        Call Ret.Add(Child)
    Loop
    
    If Ret.Count > 0 Then Set GetChildren = Ret
End Function

Private Function GetText() As String
    Dim StartIndex As Long
    StartIndex = mIndex
    
    Do
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        If mXml(mIndex) = vbLeftCheveron Then Exit Do
        mIndex = mIndex + 1
    Loop
    
    Dim Text As String
    Text = mXml.ToString(StartIndex, mIndex - StartIndex)
    Text = CorString.TrimEnd(Text, vbCrLf)
    
    mTextFormatter.Length = 0
    Call mTextFormatter.AppendString(Text)
    Call mTextFormatter.Replace(vbCr, vbLf)
    
    Dim PreviousLength As Long
    Do
        PreviousLength = mTextFormatter.Length
        Call mTextFormatter.Replace(vbLf & vbLf, vbLf)
    Loop While mTextFormatter.Length < PreviousLength
    
    Call mTextFormatter.Replace(vbLf, " ")
    GetText = mTextFormatter.ToString
End Function

Private Function GetAttributes() As Hashtable
    ' Attributes end at a '/' or '>'
    ' A '<' is an error.
    
    Dim Attributes As New Hashtable
    
    Do
        Call ConsumeWhiteSpace
        
        ' Can't end getting an attribute
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        
        Dim Name As String
        Name = GetName
        If Len(Name) = 0 Then Exit Do
        
        Call ConsumeEquals
        
        Dim Value As String
        Value = GetValue
        
        Attributes(Name) = Value
    Loop
    
    If Attributes.Count > 0 Then Set GetAttributes = Attributes
End Function

Private Sub ConsumeEquals()
    Do
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        
        Select Case mXml(mIndex)
            Case corEqualChar, corSpaceChar
            Case Else
                Exit Sub
        End Select
        
        mIndex = mIndex + 1
    Loop
End Sub

Private Function GetValue() As String
    Dim InQuotes As Boolean
    
    Dim StartIndex As Long
    StartIndex = mIndex
    
    Do
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        
        Select Case mXml(mIndex)
            Case corQuoteChar
                If Not InQuotes Then
                    InQuotes = True
                Else
                    ' Move passed the quote.
                    mIndex = mIndex + 1
                    
                    ' Don't include the quotes in the value.
                    GetValue = mXml.ToString(StartIndex + 1, mIndex - StartIndex - 2)
                    Exit Function
                End If
            
            Case corSpaceChar, corForwardSlashChar
                ' we found an unquoted value
                If Not InQuotes Then Exit Do
                
            Case vbRightCheveron
                If Not InQuotes Then
                    Exit Do
                Else
                    Throw New XmlSyntaxException
                End If
        
        End Select
        
        mIndex = mIndex + 1
    Loop

    GetValue = mXml.ToString(StartIndex, mIndex - StartIndex)
End Function

Private Function GetName() As String
    ' A name ends at a ' ', '=', '/', or '>'
    ' A '<' is always an error.
    
    Dim StartIndex As Long
    StartIndex = mIndex
    
    Do
        ' Can't end getting a name.
        If mIndex = mXml.Length Then Throw New XmlSyntaxException
        
        Select Case mXml(mIndex)
            Case corSpaceChar, corEqualChar, corForwardSlashChar, vbRightCheveron
                ' We found the end
                GetName = mXml.ToString(StartIndex, mIndex - StartIndex)
                Exit Function
            
            Case vbLeftCheveron
                Throw New XmlSyntaxException
        End Select
        
        mIndex = mIndex + 1
    Loop
End Function

Private Sub ConsumeWhiteSpace()
    Do While mIndex < mXml.Length
        Select Case mXml(mIndex)
            Case corReturnChar, corLineFeedChar, corSpaceChar, vbKeyTab
                mIndex = mIndex + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

