VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Registry Browser"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9330
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21537
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27B2
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9135
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   16113
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   16113
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This simply displays registry keys and their contents.
'
Option Explicit

''
' Displays the registry keys in the TreeView and the key-value pairs in the ListView.
'
' This will fill in a TreeView node with children if it doesn't already have them.
' Then the values for the selected subkey are displayed in the ListView.
'
Private Sub DisplaySubKeys(ByVal Node As Node, ByVal Key As RegistryKey)
    ' We assume if there are no children, then the Node hasn't
    ' had them added. If there are truly no children, then the
    ' For..Next loop will skip.
    If Node.Children = 0 Then
        ' Get all the subkey names in this key.
        Dim SubKeys() As String
        SubKeys = Key.GetSubKeyNames
        
        ' If no subkeys exist, then the array will be empty,
        ' meaning it will return -1 for the upperbound.
        Dim i As Long
        For i = 0 To UBound(SubKeys)
            ' We add the Key.Name plus the current subkey name to create a
            ' full registry key path to be associated with the node. This
            ' allows for quick access back to the registry key path.
            TreeView1.Nodes.Add Node, tvwChild, Key.Name & "\" & SubKeys(i), SubKeys(i), "Closed", "Open"
        Next i
    End If
    
    ' Now display any values also in the subkey.
    'DisplayValues Key
    
    ' Show the user what subkey we are in.
'    StatusBar1.Panels(1).Text = Key.Name
End Sub

''
' Displays all of the key-value pairs in the subkey.
'
Private Sub DisplayValues(ByVal Key As RegistryKey)
    Dim ValueKind As RegistryValueKind
    
    With ListView1
        .ListItems.Clear
        If Key Is Nothing Then Exit Sub
        
        ' We add a default entry because all subkeys contain
        ' a default, even if it doesn't exist (huh?), so just
        ' add the name to the ListView.
        Dim Item As ListItem
        Set Item = .ListItems.Add(, , "(Default)")

        ' Attempt to get the default value for the subkey.
        Dim Default As Variant
        Default = Key.GetValue("")

        ' If no default value actually exists, then the GetDefaultValue
        ' will always return an Empty value.
        If IsEmpty(Default) Then
            ' We know it doesn't exist, and Key.GetValueKind will
            ' fail, so do this manually.
            Item.SubItems(1) = "Unknown"
            Item.SubItems(2) = "(no value set)"
        Else
            ' We have a default value, so display the type of value and the value itself.
            ValueKind = Key.GetValueKind("")
            Item.SubItems(1) = ValueKindToString(ValueKind)
            Item.SubItems(2) = ValueToString(Default, ValueKind)
        End If
        
        ' Get all of names of the values in the subkey.
        Dim Values() As String
        Values = Key.GetValueNames
        
        ' Sort the names using a case insensitive string comparer.
        CorArray.Sort Values, New CaseInsensitiveComparer
        
        ' If there are no values in the subkey, then the array
        ' will be empty, meaning it will have an upperbound of -1.
        Dim i As Long
        For i = 0 To UBound(Values)
            ' An empty value name is the Default value. If no
            ' default value exists, then no empty name will exist.
            If Len(Values(i)) > 0 Then
                Set Item = .ListItems.Add(, , Values(i))
                
                ' What type of value are we getting?
                ValueKind = Key.GetValueKind(Values(i))
                
                ' Well show the type of value.
                Item.SubItems(1) = ValueKindToString(ValueKind)
                
                ' Format the value to something we can understand.
                Item.SubItems(2) = ValueToString(Key.GetValue(Values(i)), ValueKind)
            End If
        Next i
            
    End With
End Sub

''
' Formats a value to things we can understand in String format.
'
Private Function ValueToString(ByRef Value As Variant, ByVal ValueKind As RegistryValueKind) As String
    Select Case ValueKind
        Case UnknownKind:       ValueToString = "Unknown"
        Case DWordKind:         ValueToString = FormatDWord(Value)
        Case BinaryKind:        ValueToString = FormatBinary(Value)
        Case MultiStringKind:   ValueToString = Join(Value)
        Case Else:              ValueToString = Value
    End Select
End Function

''
' Formats a DWord (vbLong) to resemble the format of RegEdit.
'
Private Function FormatDWord(ByVal Value As Long) As String
    ' {0:x8}
    '      0: Reference the first argument in the argument list (Value).
    '      x: We want hex formatting, and lowercase x means lowercase formatting.
    '      8: We want the hex output to be atleast 8 characters long, so pad
    '         the beginning with zeros if necessary.
    '
    ' {0}
    '      0: Simply reference the first argument in the argument list (Value) ,
    '         and perform default formatting "G".
    '
    FormatDWord = CorString.Format("0x{0:x8} ({0})", Value)
End Function

''
' Convert the byte array into a string of hex values, each
' separated by a space.
'
Private Function FormatBinary(ByRef Bytes As Variant) As String
    Dim i As Long
    Dim sb As New StringBuilder
    
    sb.Length = 0
    For i = 0 To UBound(Bytes)
        ' {0:x2}
        '      0: Reference the first argument in the argument list (Bytes).
        '      x: We want hex formatting, and lowercase x means lowercase formatting.
        '      2: We want the hex output to be atleast 2 characters long, so pad
        '         the beginning with zeros if necessary.
        '
        ' We explicitly cast the array element to a Byte, not because it
        ' isn't a byte, but because there is a flaw in VB6 when dealing
        ' with arrays that are held in a Variant and passing the element
        ' to a ParamArray in a Class Function.
        sb.AppendFormat "{0:x2} ", CByte(Bytes(i))
    Next i
    
    ' The last character is a space. Sure it won't do anything, but
    ' this shows that you can retrieve just a substring from the StringBuilder.
    FormatBinary = sb.ToString(0, sb.Length - 1)
End Function

''
' Returns a String representation of the value kind.
'
Private Function ValueKindToString(ByVal ValueKind As RegistryValueKind) As String
    Select Case ValueKind
        Case UnknownKind:       ValueKindToString = "Unknown"
        Case StringKind:        ValueKindToString = "REG_SZ"
        Case DWordKind:         ValueKindToString = "REG_DWORD"
        Case BinaryKind:        ValueKindToString = "REG_BINARY"
        Case MultiStringKind:   ValueKindToString = "REG_MULTI_SZ"
        Case ExpandStringKind:  ValueKindToString = "REG_EXPAND_SZ"
        Case QWordKind:         ValueKindToString = "REG_QWORD"
    End Select
End Function

''
' Shows the available root nodes to begin browsing from.
'
Private Sub FillRootKeys()
    Dim Node As Node
    Set Node = TreeView1.Nodes.Add(, , "Computer", "My Computer", "Closed", "Open")
    TreeView1.Nodes.Add Node, tvwChild, Registry.ClassesRoot.Name, Registry.ClassesRoot.Name, "Closed", "Open"
    TreeView1.Nodes.Add Node, tvwChild, Registry.CurrentConfig.Name, Registry.CurrentConfig.Name, "Closed", "Open"
    TreeView1.Nodes.Add Node, tvwChild, Registry.CurrentUser.Name, Registry.CurrentUser.Name, "Closed", "Open"
    TreeView1.Nodes.Add Node, tvwChild, Registry.LocalMachine.Name, Registry.LocalMachine.Name, "Closed", "Open"
End Sub

''
' Gets things going.
Private Sub Form_Load()
    With ListView1
        .ColumnHeaders.Add , , "Name", 2500
        .ColumnHeaders.Add , , "Type", 1500
        .ColumnHeaders.Add , , "Data", 2500
    End With
    
    FillRootKeys
End Sub

''
' Displays a key that was clicked on.
'
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Key As RegistryKey
    
    Me.MousePointer = ccHourglass
    
    Select Case Node.Key
        Case "Computer"
            ' do nothing
            
        ' If one of the Root keys was clicked, we pass in
        ' the specific root key object.
        Case Registry.ClassesRoot.Name:     Set Key = Registry.ClassesRoot
        Case Registry.CurrentConfig.Name:   Set Key = Registry.CurrentConfig
        Case Registry.CurrentUser.Name:     Set Key = Registry.CurrentUser
        Case Registry.LocalMachine.Name:    Set Key = Registry.LocalMachine
    
        ' As we added nodes, we stored the name of the subkey
        ' as the key in the node. This allows easy access back
        ' to the registry key through the full path name.
        Case Else
            Dim Parts() As String
            
            Parts = CorString.Split(Node.Key, "\", 2)
            
            Select Case Parts(0)
                Case Registry.ClassesRoot.Name: Set Key = Registry.ClassesRoot.OpenSubKey(Parts(1))
                Case Registry.CurrentConfig.Name: Set Key = Registry.CurrentConfig.OpenSubKey(Parts(1))
                Case Registry.CurrentUser.Name: Set Key = Registry.CurrentUser.OpenSubKey(Parts(1))
                Case Registry.LocalMachine.Name: Set Key = Registry.LocalMachine.OpenSubKey(Parts(1))
            End Select
    End Select

    
    DisplaySubKeys Node, Key
    DisplayValues Key
    Me.MousePointer = ccDefault
    
    If Not Key Is Nothing Then StatusBar1.Panels(1).Text = Key.Name
End Sub
