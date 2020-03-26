Attribute VB_Name = "modMain"
' This example simply increases the height of the console
' window and draws a colors square in the center. The user
' is then prompted to hit a key.
'
' *** WARNING ***
' Remember to never click on the Console close button, or
' press the End button in the IDE, or the application will crash.
'
Option Explicit

Private Sub Main()
    ' Set our console title.
    Console.Title = "A Colored Square!!"
    
    ' Lets double the height of the console window.
    Console.SetWindowSize Console.WindowWidth, Console.WindowHeight * 2
    
    ' And get rid of the cursor.
    Console.CursorVisible = False
    
    ' Find the center of the console window width.
    Dim HalfWidth As Long
    HalfWidth = Console.WindowWidth \ 2
    
    ' Find the center of the console window height.
    Dim HalfHeight As Long
    HalfHeight = Console.WindowHeight \ 2
    
    ' Lets define the outside border of the square we
    ' are going to draw. Lets make it 10 characters
    ' in each direction from the center.
    Dim RightSide As Long
    RightSide = HalfWidth + 10
    
    Dim BottomSide As Long
    BottomSide = HalfHeight + 10
    
    Dim LeftSide As Long
    LeftSide = HalfWidth - 10
    
    Dim TopSide As Long
    TopSide = HalfHeight - 10
    
    ' Start our drawing position in the upper-left
    ' corner of the border we've defined.
    Dim x As Long
    x = LeftSide
    
    Dim y As Long
    y = TopSide
    
    Dim Color As ConsoleColor
    Do
        SetColor
        
        ' Draw a line of characters from left to right.
        Do While x < RightSide
            Console.SetCursorPosition x, y
            Console.WriteValue "*"
            x = x + 1
        Loop
        
        SetColor
        
        ' Draw a line of characters down the right side
        ' of the square border.
        Do While y < BottomSide
            Console.SetCursorPosition x, y
            Console.WriteValue "*"
            y = y + 1
        Loop
        
        SetColor
        
        ' Draw a line across the bottom, back towards the left.
        Do While x > LeftSide
            Console.SetCursorPosition x, y
            Console.WriteValue "*"
            x = x - 1
        Loop
        
        SetColor
        
        ' Draw a line up the left side of the square border.
        Do While y > TopSide
            Console.SetCursorPosition x, y
            Console.WriteValue "*"
            y = y - 1
        Loop
    
        ' Now shrink the borders in by one character from all sides.
        RightSide = RightSide - 1
        LeftSide = LeftSide + 1
        TopSide = TopSide + 1
        BottomSide = BottomSide - 1
        
        ' We'll keep going until the left and right side touch.
    Loop While LeftSide <= RightSide
    
    ' Move the cursor so our text will be vertically centered,
    ' and a couple rows below the colored square.
    Console.SetCursorPosition HalfWidth - (Len("Press Any Key") / 2), HalfHeight + 12
    
    ' Wait for the user to obey.
    Console.WriteLine "Press Any Key"
    Console.ReadKey
End Sub

' This simply sets the fore color of the characters being
' printed to the console. The Color variable is incremented,
' and if it becomes black, we skip it, to keep everything visible.
Private Sub SetColor()
    Static Color As ConsoleColor
    
    Color = (Color + 1) Mod 16
    If Color = Black Then SetColor
    Console.ForegroundColor = Color
End Sub
