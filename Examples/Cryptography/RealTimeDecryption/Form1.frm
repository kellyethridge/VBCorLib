VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line Reader"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReadLine 
      Caption         =   "Read Line"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtCipherText 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtPlainText 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Demonstrates how an encrypted data source can read from or written to
' in the same manner as a plain text data source.
'
' In this case we are dealing with files. We have two files.
'
' One Contains 10 lines of plain text. Something that can easily be created
' and read from in a dynamic fashion. Meaning that we don't need to read the
' whole file into memory just to display the lines contained within.
'
' The other contains 10 lines as well, but the file itself is completely
' encrypted. Normal conventions would require the entire file to be loaded
' and decrypted in order to obtain the data within. Using a CryptoStream
' we show that an encrypted file an be read from in a dynamic manner as a
' plain text file. The CryptoStream will decrypt the file on the fly as
' is needed by the reading mechanism, in this case a StreamReader object.
'
' Write to an encrypted file works the same way. This is demonstrated in
' the CreateEncryptedFile method. This shows how to setup a file to be
' written to and have the data encrypted as it is written to the file,
' on the fly of course.
'
Option Explicit

' Used to read text lines from the plain text file.
Private mTextReader As StreamReader

' Used to read text lines from the encrypted text file.
' Each line is decrypted on the fly as needed.
Private mEncryptedReader As StreamReader

' The cipher to be used for encrypting and decrypting our file.
Private mCipher As New RijndaelManaged

Private mPlainTextFinished As Boolean
Private mEncryptedFinished As Boolean


Private Sub cmdReadLine_Click()
    If Not mPlainTextFinished Then mPlainTextFinished = ReadLine(mTextReader, txtPlainText)
    If Not mEncryptedFinished Then mEncryptedFinished = ReadLine(mEncryptedReader, txtCipherText)
End Sub

Private Sub Form_Load()
    Call CreateSourceFiles
    Call OpenSourceFiles
End Sub


' We use the same code to read from both streams.
Private Function ReadLine(ByVal Reader As StreamReader, ByVal Box As TextBox) As Boolean
    Dim Line As String
    Line = Reader.ReadLine
    
    If Not CorString.IsNull(Line) Then
        Box.Text = Box.Text & Line & vbCrLf
    Else
        Call Reader.CloseReader
        Box.Text = Box.Text & "Finished"
        ReadLine = True
    End If
End Function


' Open the two files for reading.
Private Sub OpenSourceFiles()
    ' Opening the plain text file is simple.
    Set mTextReader = NewStreamReader(Path.Combine(App.Path, "PlainTextLines.txt"))

    ' Opening the encrypted file takes some extra steps.
    ' We open our FileStream like normal.
    Dim InputStream As FileStream
    Set InputStream = File.OpenFile(Path.Combine(App.Path, "EncryptedTextLines.txt"), FileMode.OpenExisting)
    
    ' Now we need to wrap the FileStream with our decrypting CryptoStream.
    ' This will perform the decryption on the fly as needed. It does not
    ' need to read in the entire file at once. Only what is needed as each
    ' line is read by the StreamReader.
    Dim DecryptStream As CryptoStream
    Set DecryptStream = NewCryptoStream(InputStream, mCipher.CreateDecryptor, CryptoStreamMode.ReadMode)
    
    ' Place our CryptoStream in a StreamReader that lets up deal with the
    ' stream as text, instead of as a bunch of bytes.
    Set mEncryptedReader = NewStreamReader(DecryptStream)
End Sub

Private Sub CreateSourceFiles()
    Call CreatePlainTextFile
    Call CreateEncryptedFile
End Sub

' Ensure the plain text file exists.
'
' This is pretty straight forward. We will open a file stream to be written
' to using a StreamWriter later on. There is nothing that needs to be done
' with the data as it is being written.
Private Sub CreatePlainTextFile()
    ' Build our filename.
    Dim PlainTextFile As String
    PlainTextFile = Path.Combine(App.Path, "PlainTextLines.txt")
    
    ' We open a stream to the file to be written to. This is the same
    ' as writing bytes directly to the byte without any conversion or
    ' encoding being applied. This is all binary access.
    Dim OutputStream As FileStream
    Set OutputStream = File.OpenFile(PlainTextFile, FileMode.Create)
    
    ' Have the lines be written to the stream.
    Call WriteLines("Plain Text Line", OutputStream)
End Sub

' Ensure the encrypted file exists.
'
' This setup is a little more complex than the simple plain text file version.
' We need to create the file stream, same as the plain text version, but then
' wrap that stream in a CryptoStream to an encryption transform can be applied
' while data is being written.
'
Private Sub CreateEncryptedFile()
    ' Build our filename.
    Dim EncryptedTextFile As String
    EncryptedTextFile = Path.Combine(App.Path, "EncryptedTextLines.txt")
    
    ' We open a stream to the file to be written to. This is the same
    ' as writing bytes directly to the byte without any conversion or
    ' encoding being applied. This is all binary access.
    Dim OutputStream As FileStream
    Set OutputStream = File.OpenFile(EncryptedTextFile, FileMode.Create)
    
    ' To write our lines out in an encrypted form, we create a second
    ' stream that performs the encryption on the fly while, then the
    ' encrypted data will be written to the OutputStream stream.
    Dim EncryptStream As CryptoStream
    
    ' Pass in the OutputStream to be written to by this CryptoStream object,
    ' also, pass in the Tranform that will be applied to data as it is
    ' written to this CryptoStream object.
    Set EncryptStream = NewCryptoStream(OutputStream, mCipher.CreateEncryptor, CryptoStreamMode.WriteMode)
    
    ' Pass the CryptoStream object to the WriteLines function as a normal
    ' Stream object. The function doesn't care what kind of Stream object
    ' it is dealing with, just as long as it implements the Stream interface.
    Call WriteLines("Encrypted Text Line", EncryptStream)
End Sub

' This will write 10 text lines then close the writer, which will close the stream.
Private Sub WriteLines(ByVal Phrase As String, ByVal Stream As Stream)
    ' Create a writer for the stream object. This will make things
    ' much easier with writing text to an underlying stream.
    Dim Writer As StreamWriter
    Set Writer = NewStreamWriter(Stream)
    
    Dim i As Long
    For i = 1 To 10
        Call Writer.WriteLine(CorString.Format("{0} {1}", Phrase, i))
    Next i
    
    Call Writer.CloseWriter
End Sub
