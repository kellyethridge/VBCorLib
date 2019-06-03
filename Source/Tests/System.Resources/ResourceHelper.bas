Attribute VB_Name = "ResourceHelper"
Option Explicit

Public Function LoadCursor() As StdPicture
    Set LoadCursor = LoadPicture("normal01.cur")
End Function

Public Function LoadIcon() As StdPicture
    Set LoadIcon = LoadPicture("checkmrk.ico")
End Function

Public Function LoadBitMap() As StdPicture
    Set LoadBitMap = LoadPicture("balloon.bmp")
End Function

Private Function LoadPicture(ByVal FileName As String) As StdPicture
    Set LoadPicture = VB.LoadPicture(MakeResourcePath(FileName))
End Function

Public Function MakeResourcePath(ByVal FileName As String) As String
    MakeResourcePath = Path.Combine(App.Path, CorString.Format("System.Resources/{0}", FileName))
End Function
