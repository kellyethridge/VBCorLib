VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Encryption"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   7455
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "Decrypt"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtInputFileName 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton cmbBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "Encrypt"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtOutputFileName 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   5535
      End
      Begin MSComctlLib.ProgressBar pbrProgress 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Input:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Output:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboBlockSizes 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cboPadding 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4560
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "Form1.frx":003F
      Left            =   4560
      List            =   "Form1.frx":0052
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cboKeySizes 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cboAlgorithms 
      Height          =   315
      ItemData        =   "Form1.frx":006F
      Left            =   1200
      List            =   "Form1.frx":007F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Block Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Padding:"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Mode:"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Key Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Algorithm:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is a simple demonstration on how to encrypt and decrypt a file
' using any of the SymmetricalAlgorithm implementations and the CryptoStream.
'
' Also, the Rfc2898DeriveBytes class is demonstrated.
'
Option Explicit

Private mCsp As SymmetricAlgorithm



Private Sub cboAlgorithms_Click()
    ' We can create cryptographic objects based on names using the CryptoConfig class.
    Set mCsp = CryptoConfig.CreateFromName(cboAlgorithms.Text)
    
    ' Update all the ComboBoxes
    Call FillKeySizesComboBox
    Call FillBlockSizesComboBox
End Sub

Private Sub cmbBrowse_Click()
    On Error GoTo errTrap
    With CD
        .CancelError = True
        .FileName = txtInputFileName.Text
        .ShowOpen
        txtInputFileName.Text = .FileName
        txtOutputFileName.Text = .FileName & ".Out"
    End With
errTrap:
End Sub

Private Sub cmdDecrypt_Click()
    On Error GoTo errTrap
    
    ' We need to prepare the Csp before getting the
    ' Transform so it will reflect the appropriate settings.
    Call PrepareCsp
    
    ' Pass in the Transform to be used on the output stream.
    Call DoCipher(mCsp.CreateDecryptor)
    Exit Sub
    
errTrap:
    Dim Ex As Exception
    If Catch(Ex, Err) Then Call MsgBox(Ex.ToString)
End Sub

Private Sub cmdEncrypt_Click()
    On Error GoTo errTrap
    
    ' We need to prepare the Csp before getting the
    ' Transform so it will reflect the appropriate settings.
    Call PrepareCsp
    
    ' Pass in the Transform to be used on the output stream.
    Call DoCipher(mCsp.CreateEncryptor)
    Exit Sub
    
errTrap:
    Dim Ex As Exception
    If Catch(Ex, Err) Then Call MsgBox(Ex.ToString)
End Sub

Private Sub Form_Load()
    cboPadding.ListIndex = 1
    cboMode.ListIndex = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoCipher(ByVal t As ICryptoTransform)
    ' Open our input stream. This can be a file to be encrypted or decrypted.
    Dim InStream As Stream
    Set InStream = File.OpenFile(txtInputFileName.Text, FileMode.OpenExisting)
    
    ' Open the output stream. This will be a file that is encrypted or decrypted
    ' based on the type of ICryptoTransform passed in. The output stream will
    ' process the output data using the supplied transform.
    Dim OutStream As Stream
    Set OutStream = Cor.NewCryptoStream(File.OpenFile(txtOutputFileName.Text, FileMode.Create), t, CryptoStreamMode.WriteMode)
    
    ' Let'er rip!
    Call ProcessFile(InStream, OutStream)
    Call OutStream.CloseStream
    Call InStream.CloseStream
End Sub

''
' Once everything has been selected by the user, we need to
' change the settings of our service provider.
Private Sub PrepareCsp()
    mCsp.Mode = cboMode.ListIndex + 1
    mCsp.Padding = cboPadding.ListIndex + 1
    mCsp.KeySize = cboKeySizes.ItemData(cboKeySizes.ListIndex)
    mCsp.BlockSize = cboBlockSizes.ItemData(cboBlockSizes.ListIndex)
    
    ' Generate our secret key based on our text password.
    mCsp.Key = GenerateKey(txtPassword.Text, mCsp.KeySize \ 8)
    
    ' We will generate an array of bytes with values of zero to keep things simple.
    mCsp.IV = CorArray.CreateInstance(vbByte, mCsp.BlockSize \ 8)
End Sub

''
' Process a file. This might be encrypting or decrypting. The process is
' the same either way. The streams handle any ciphering that happens.
'
' The slowest part of this routine is the ProgressBar update. A better
' system would only update the ProgressBar when it should actually change,
' however, this is a simple demonstration.
Private Sub ProcessFile(ByVal Src As Stream, ByVal Dst As Stream)
    pbrProgress.Max = Src.Length
    pbrProgress.Value = 0
    
    ' We will use a 4K buffer.
    ReDim b(4095) As Byte
    Dim BytesRead As Long
    
    ' Read the first block in. We don't care what kind of
    ' stream we are reading from. We just want the bytes.
    BytesRead = Src.ReadBlock(b, 0, 4096)
    
    Do While BytesRead > 0
        ' Perform the slow part.
        pbrProgress.Value = pbrProgress.Value + BytesRead
        
        ' We know that all encryption and decryption will happen
        ' through this stream. The data will be transformed using
        ' the ICryptoTransform that was passed to the CryptoStream.
        ' However, we treat the output stream just like any stream.
        Call Dst.WriteBlock(b, 0, BytesRead)
        
        ' Read in the next block of bytes.
        BytesRead = Src.ReadBlock(b, 0, 4096)
    Loop
End Sub

Private Sub FillKeySizesComboBox()
    Call FillSizes(mCsp.LegalKeySizes, cboKeySizes)
    cboKeySizes.ListIndex = 0
End Sub

Private Sub FillBlockSizesComboBox()
    Call FillSizes(mCsp.LegalBlockSizes, cboBlockSizes)
    cboBlockSizes.ListIndex = 0
End Sub

''
' Fills a ComboBox with KeySizes values.
'
Private Sub FillSizes(ByRef KeySizes() As KeySizes, ByVal Box As ComboBox)
    Call Box.Clear
    
    ' We know that there is only 1 element in
    ' the KeySizes array, so take advantage of that.
    Dim Sizes As KeySizes
    Set Sizes = KeySizes(0)
    
    ' If a SkipSize is zero, then the For loop will never
    ' advance, so we'd be stuck forever. So we just handle
    ' this special case right here. The MinSize and MaxSize
    ' will be the same for this type of case.
    If Sizes.SkipSize = 0 Then
        Call Box.AddItem(Sizes.MinSize & " bits")
        Box.ItemData(Box.NewIndex) = Sizes.MinSize
    Else
        ' This will fill a ComboBox with all the possible
        ' key size values computed from the KeySizes object.
        Dim i As Long
        For i = Sizes.MinSize To Sizes.MaxSize Step Sizes.SkipSize
            Call Box.AddItem(i & " bits")
            Box.ItemData(Box.NewIndex) = i
        Next i
    End If
End Sub

''
' Generates a byte array key from a text password.
'
Private Function GenerateKey(ByVal Password As String, ByVal KeySize As Long) As Byte()
    ' We will have a salt of bytes with values of zero to keep things simple.
    Dim Salt() As Byte
    ReDim Salt(0 To KeySize - 1)
    
    ' Use the recommended key generator. This will generate an unlimited
    ' number of bytes from a single text password or byte array. If the
    ' generator is reset, then the same sequence of bytes will be produced.
    Dim Generator As Rfc2898DeriveBytes
    Set Generator = Cor.NewRfc2898DeriveBytes(Password, Salt)
    
    ' Return the number of bytes needed for the specific
    ' encryption algorithm key size. This is in bytes, not bits.
    GenerateKey = Generator.GetBytes(KeySize)
End Function

