VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Very Simple File Browser (Uses a Custom IComparer class for sorting)"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   16113
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   16113
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
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
Option Explicit

Private mSortComparer As FileComparer


''
' Simply add new nodes to the treeview if needed.
'
Private Sub ShowSubDirectories(ByVal Node As Node)
    ' Create a DirectoryInfo object using the full path
    ' supplied by the TreeView. This happens to be the
    ' same path structure as a directory structure.
    '
    ' Remember, we want a directory name, so don't include
    ' an asterisc '*' at the end.
    Dim Dir As DirectoryInfo
    Set Dir = NewDirectoryInfo(Node.FullPath)
    
    ' Lets ensure that we can access the specific drive.
    Dim Drive As DriveInfo
    Set Drive = NewDriveInfo(Dir.Root.Name)
    If Not Drive.IsReady Then
        MsgBox CorString.Format("Drive '{0}' is not ready.", Drive.Name)
        Exit Sub
    End If
    
    ' If there are no children, then we probably haven't filled
    ' it with the subdirectory names, so do it this time.
    If Node.Children = 0 Then
        ' Fetch the subdirectories in the main directory just created.
        Dim Subs() As DirectoryInfo
        Subs = Dir.GetDirectories
        
        ' Iterate through the subdirectories, creating a new node for each.
        With TreeView1
            Dim i As Long
            For i = 0 To UBound(Subs)
                .Nodes.Add Node, tvwChild, , Subs(i).Name, "Closed", "Open"
            Next i
        End With
    End If
    
    ' now show the contents of the selected directory in the
    ' panel next to the treeview.
    ShowFiles Dir
End Sub

''
' Show the files of the specified directory.
'
Private Sub ShowFiles(ByVal SubDir As DirectoryInfo)
    ' Retrieve all of the files in the subdirectory.
    Dim Files() As FileInfo
    Files = SubDir.GetFiles
    
    ' Sort the files using our custom comparer.
    CorArray.Sort Files, mSortComparer
    
    ' Add each file to the listview, including the additional
    ' information supplied for each column in the listview.
    With ListView1
        .ListItems.Clear
        
        Dim i As Long
        For i = 0 To UBound(Files)
            Dim Item As ListItem
            Set Item = .ListItems.Add(, , Files(i).Name)
            
            ' This formats the file length by creating a column that
            ' is 10 characters wide and right aligned by specifying the 10 to
            ' create columns. A negative value (-10) would create a left aligned column.
            ' The 'N' indicates we want Number formatting which will include
            ' group separators (commas) in the number.
            ' A zero is appended to 'N' to prevent any decimal places to be included.
            '
            Item.SubItems(1) = CorString.Format("{0,10:N0} KB", Ceiling(Files(i).Length / 1000))
            
            ' LastAccessTime returns a Variant containing a
            ' cDateTime object. The reason is to allow the Let
            ' property to accept either a cDateTime object or
            ' a VB Date value when setting the property.
            '
            ' Since a cDateTime object is always returned, call
            ' its ToString function to get the current date and time.
            '
            Item.SubItems(2) = Files(i).LastAccessTime.ToString
        Next i
    End With
End Sub

''
' Displays a list of drives in the TreeView.
'
Private Sub ShowDrives()
    ' Retrieve all the known drives.
    Dim Drives() As String
    Drives = Environment.GetLogicalDrives
    
    ' Add each drive to the TreeView.
    Dim i As Long
    For i = 0 To UBound(Drives)
        TreeView1.Nodes.Add , , , Drives(i), "Closed", "Open"
    Next i
    TreeView1.Nodes(1).Selected = True
End Sub

''
' Initialize everything.
'
Private Sub Form_Load()
    ' Setup our file comparer for sorting.
    ' Yes, the listview can sort, but this
    ' is to demonstrate other sorting methods.
    Set mSortComparer = New FileComparer
    mSortComparer.SortColumn = "Name"
    mSortComparer.SortOrder = Ascending
    
    ' Create a simple listview for the files.
    With ListView1
        .ColumnHeaders.Add , "Name", "Name", 2500
        .ColumnHeaders.Add , "Size", "Size", , lvwColumnRight
        .ColumnHeaders.Add , "Modified", "Date Modified", 2500
    End With
    
    ShowDrives
    ShowSubDirectories TreeView1.Nodes(1)
End Sub

''
' Change sorting parameters.
'
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ColumnHeader
        ' if the column clicked is the same that we already sort
        ' by, then we must be changing the sort order on that column,
        ' so toggle the sort type in our custom comparer.
        If .Key = mSortComparer.SortColumn Then
            mSortComparer.SortOrder = 1 - mSortComparer.SortOrder
        Else
            ' A different column has been clicked than what we are sorting
            ' by, so set the new sort column in our custom comparer.
            mSortComparer.SortColumn = .Key
        End If
        
        ' We will assume a node is still selected, so display it.
        ShowSubDirectories TreeView1.SelectedItem
    End With
End Sub

''
' Select a node to display.
'
Private Sub TreeView1_Click()
    ShowSubDirectories TreeView1.SelectedItem
End Sub
