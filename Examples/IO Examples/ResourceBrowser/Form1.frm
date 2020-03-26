VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource File Browser"
   ClientHeight    =   8670
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   6480
      ScaleHeight     =   3555
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   14843
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenRes 
         Caption         =   "&Open RES"
      End
      Begin VB.Menu mnuFileOpenEXE 
         Caption         =   "&Open EXE"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mResources As New Hashtable



Private Sub LoadResources(ByVal Reader As IResourceReader)
    Call ClearException
    On Error GoTo errTrap   ' will throw an exception for invalid file formats.
    
    Call mResources.Clear
    Call ListView1.ListItems.Clear
    
    ' Iterate all the loaded resources.
    Dim Entry As DictionaryEntry
    For Each Entry In Reader
        ' We need to access ResourceKey functions.
        Dim Key As ResourceKey
        Set Key = Entry.Key
            
        ' Store the resource locally using the key.
        Call mResources.Add(Key.ToString, Entry.Value)
        
        Dim Item As ListItem
        
        ' Use the Key.ToString to associate the resource
        ' with the listview item.
        Set Item = ListView1.ListItems.Add(, Key.ToString, Key.ResourceName)
        Item.SubItems(1) = GetTypeName(Key.ResourceType)
        Item.SubItems(2) = Key.LanguageID
    Next Entry
    
    Reader.CloseReader
    
    Exit Sub
    
errTrap:
    Dim ex As Exception
    If Catch(ex) Then
        MsgBox ex.ToString, vbOKOnly + vbExclamation, "Error"
    End If
End Sub

Private Sub ShowResource(ByVal Key As String)
    Dim Value As Variant
    
    ' Use the VBCorLib supplied convenient function.
    Call MoveVariant(Value, mResources(Key))
    
    Call Picture1.Cls
    
    ' We have to perform the object check this
    ' way because the StdPicture has a Default function,
    ' so the VarType ends up getting the Default function's
    ' return value and determining the type from that instead.
    If IsObject(Value) Then
        If TypeOf Value Is StdPicture Then
            ' If it's a picture, then draw it.
            Call Picture1.PaintPicture(Value, 0, 0)
            
        ElseIf TypeOf Value Is PictureResourceGroup Then
            ' If it's a group, then display the
            ' details of each group entry.
            Dim Group As PictureResourceGroup
            Set Group = Value
            
            Dim i As Long
            For i = 0 To Group.Count - 1
                Picture1.Print GetGroupInfo(Group(i), Group.GroupType)
            Next i
        End If
    Else
        Select Case VarType(Value)
            Case vbString
                Picture1.Print Value
            Case Else
                Picture1.Print "Unknown"
        End Select
    End If
End Sub

''
' Returns a formatted group entry.
'
Private Function GetGroupInfo(ByVal Info As PictureResourceInfo, ByVal GroupType As PictureGroupTypes) As String
    GetGroupInfo = CorString.Format("{0} Resource ID: {1}, Size: {2}x{3}, Colors: {4}", IIf(GroupType = IconGroup, "Icon", "Cursor"), Info.ResourceID, Info.Width, Info.Height, Info.Colors)
End Function

Private Function GetTypeName(ByRef ResourceType As Variant) As String
    If VarType(ResourceType) = vbString Then
        GetTypeName = ResourceType
    Else
        Select Case CLng(ResourceType)
            Case CursorResource:    GetTypeName = "Cursor"
            Case BitmapResource:    GetTypeName = "Bitmap"
            Case IconResource:      GetTypeName = "Icon"
            Case MenuResource:      GetTypeName = "Menu"
            Case DialogBox:         GetTypeName = "Dialog Box"
            Case stringresource:    GetTypeName = "String"
            Case FontDirectory:     GetTypeName = "Font Directory"
            Case FontResource:      GetTypeName = "Font"
            Case AcceleratorTable:  GetTypeName = "Accelerator Table"
            Case UserDefined:       GetTypeName = "User Defined"
            Case GroupCursor:       GetTypeName = "Cursor Group"
            Case GroupIcon:         GetTypeName = "Icon Group"
            Case VersionResource:   GetTypeName = "Version"
            Case DialogInclude:     GetTypeName = "Dialog Include"
            Case PlugPlay:          GetTypeName = "Plug And Play"
            Case VXD:               GetTypeName = "VXD"
            Case AniCursor:         GetTypeName = "Animated Cursor"
            Case AniIcon:           GetTypeName = "Animated Icon"
            Case HTML:              GetTypeName = "HTML"
            Case Else:              GetTypeName = "Unknown"
        End Select
    End If
End Function

Private Sub Init()
    With ListView1
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            Call .Add(, , "Name")
            Call .Add(, , "Type")
            Call .Add(, , "Language")
        End With
    End With
End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub ListView1_Click()
    Call ShowResource(ListView1.SelectedItem.Key)
End Sub

Private Sub mnuFileOpenEXE_Click()
    Dim FileName As String
    FileName = GetFileName("Executable (*.EXE)|*.EXE|Library (*.DLL)|*.DLL|User Control (*.OCX)|*.OCX")
    If Len(FileName) > 0 Then
        Call LoadResources(Cor.NewWinResourceReader(FileName))
    End If
End Sub

Private Sub mnuFileOpenRes_Click()
    Dim FileName As String
    FileName = GetFileName("Resource (*.RES)|*.RES")
    If Len(FileName) > 0 Then
        Call LoadResources(Cor.NewResourceReader(FileName))
    End If
End Sub

Private Function GetFileName(ByVal Filter As String) As String
    On Error GoTo errTrap
    With CD
        .CancelError = True
        .DialogTitle = "Open Resource File."
        .Filter = Filter
        Call .ShowOpen
        GetFileName = .FileName
    End With
errTrap:
End Function
