VERSION 5.00
Begin VB.Form Wand 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   6270
   ControlBox      =   0   'False
   DrawWidth       =   5
   FillColor       =   &H00FF00FF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "Wand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSpeed 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Left            =   120
      ScaleHeight     =   1140
      ScaleWidth      =   5370
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5370
      Begin VB.HScrollBar hscSpeed 
         Height          =   570
         LargeChange     =   20
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   1
         Top             =   120
         Value           =   60
         Width           =   5130
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fast                    Medium                    Slow"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   5130
      End
   End
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5700
      Top             =   120
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5700
      Top             =   840
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuControls 
         Caption         =   "&Controls"
         Begin VB.Menu mnuDeskTop 
            Caption         =   "Show &DeskTop WallPaper"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "&Speed"
         End
         Begin VB.Menu mnuTrail 
            Caption         =   "&Trail Off"
         End
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Wand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the "Magic Wand" program.                            ****
'   ****                                                                 ****
'   ****    Written by:  Randy Giese                                     ****
'   ****    Written at:  www.RandyGrams.com                              ****
'   ****    Written on:  January 06, 2005                                ****
'   ****    E-Mail me:   WebMaster@RandyGrams.com                        ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    The Wand flies around the screen bouncing off the walls.     ****
'   ****    It is constantly changing colors to create a Gradient        ****
'   ****    effect.                                                      ****
'   ****                                                                 ****
'   ****    Click the Right Mouse Button for the following options.      ****
'   ****                                                                 ****
'   ****    The first item on the Menu is Trail.                         ****
'   ****    If you turn the Trail off, the Wand still flies around       ****
'   ****    the screen and bounces off the walls, but it doesn't         ****
'   ****    leave a Gradient Trail behind.                               ****
'   ****                                                                 ****
'   ****    The other option is Speed.                                   ****
'   ****    Click it to open the Speed control.                          ****
'   ****                                                                 ****
'   ****    Double-Click anywhere on the form to end the Magic Wand.     ****
'   ****                                                                 ****
'   ****    That's it.  Have fun and enjoy using the Magic Wand.         ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
Option Explicit

Private blnTrail                As Boolean
Private blnWand                 As Boolean

Private curEnd                  As Currency
Private curFrequency            As Currency
Private curStart                As Currency

Private lngAngle(1 To 4)        As Long
Private lngCenterX              As Long
Private lngCenterY              As Long
Private lngDir(1 To 4)          As Long
Private lngPorM(3)              As Long
Private lngRGB(3)               As Long
Private lngSpeed                As Long
Private lngXY(1 To 8)           As Long

'   Define the constants for GetQueueStatus
Private Const QS_HOTKEY         As Long = &H80
Private Const QS_KEY            As Long = &H1
Private Const QS_MOUSEBUTTON    As Long = &H4
Private Const QS_MOUSEMOVE      As Long = &H2
Private Const QS_PAINT          As Long = &H20
Private Const QS_POSTMESSAGE    As Long = &H8
Private Const QS_SENDMESSAGE    As Long = &H40
Private Const QS_TIMER          As Long = &H10
Private Const QS_ALLINPUT       As Long = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

'   Declarations
'Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
'Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
'Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
'Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'   Declarations
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Sub Form_Load()

    RG_About.Move (Screen.Width - RG_About.Width) \ 2, (Screen.Height - RG_About.Height) \ 2
    RG_Help.Help_File = App.Path & "\Wand.HTML"

    blnTrail = True
    Randomize Timer          '  Variety is the spice of life
    Me.Show

'   Locate the Center of the Screen
    lngCenterX = Screen.Width \ Screen.TwipsPerPixelX \ 2
    lngCenterY = Screen.Height \ Screen.TwipsPerPixelY \ 2

'   Move the Mouse Pointer to the Center of the Screen
    SetCursorPos lngCenterX, lngCenterY
'   Then hide the Mouse Pointer
    MouseIcon = LoadPicture(App.Path & "\Blank.Ico")
    MousePointer = 99

    subCurFreq
    subNewWand

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'(Button As Integer, Shift As Integer, X As Single, Y As Single)

'   Check to see if the Mouse really did move.
    If Abs(x - lngCenterX) > 1 Then
'       If it did, save the new Pointer location and show it.
        lngCenterX = x
        subShowMouse
    End If

'   Check to see if the Mouse really did move.
    If Abs(y - lngCenterY) > 1 Then
'       If it did, save the new Pointer location and show it.
        lngCenterY = y
        subShowMouse
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
'        PopupMenu mnuMenu
        mnuMenu.Visible = True
'       Move the Mouse Pointer to the Menu.
        SetCursorPos 10, 10
    Else
        subNewWand
    End If

End Sub

Private Sub Form_Paint()

    If mnuDeskTop.Checked = True Then
        PaintDesktop Me.hdc
    Else
        Cls
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    RG_Quit

End Sub

Private Sub hscSpeed_Change()

    lngSpeed = hscSpeed.Value / 10

End Sub

Private Sub mnuAbout_Click()
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the 'About' screen.                                  ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    mnuMenu.Visible = False
    RG_About.Show vbModal, Me

End Sub

Private Sub mnuDeskTop_Click()

    mnuDeskTop.Checked = Not mnuDeskTop.Checked
    subNewWand

End Sub

Private Sub mnuHelp_Click()
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the 'Help' screen.                                   ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    mnuMenu.Visible = False
    RG_Help.Show vbModal, Me

End Sub

Private Sub mnuQuit_Click()

    RG_Quit

End Sub

Private Sub mnuSpeed_Click()

    mnuMenu.Visible = False
    picSpeed.Visible = True
    tmrSpeed.Enabled = True

End Sub

Private Sub mnuTrail_Click()

    If mnuTrail.Caption = "&Trail Off" Then
        blnTrail = False
        mnuTrail.Caption = "&Trail On"
    Else
        blnTrail = True
        mnuTrail.Caption = "&Trail Off"
    End If

    mnuMenu.Visible = False
    Cls

End Sub

Private Sub RG_Quit()

Dim lngCntr                           As Long
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is my standard Quit/End routine.  I make sure the       ****
'   ****    Mouse Pointer is turned on, close all open files, then       ****
'   ****    close all forms and End the program.                         ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    On Error Resume Next

    DoEvents
    Close

    For lngCntr = (Forms.Count - 1) To 0 Step -1
        Unload Forms(lngCntr)               '  Unloads ALL forms
    Next lngCntr

    End

    On Error GoTo 0

End Sub

Private Sub subChgColors()

Dim lngCntr                           As Long

'   This updates the Gradient's colors
    lngRGB(1) = lngRGB(1) + lngPorM(1)      '  Change first color value

'   Check all three color groups to make sure they are within legal bounds.
    For lngCntr = 1 To 3
        If lngRGB(lngCntr) < 1 Or lngRGB(lngCntr) > 254 Then              '  If color reaches min or max,
            lngPorM(lngCntr) = -lngPorM(lngCntr)           '  Reverse the color direction

            If lngCntr < 3 Then             '  If it's not the last color group (Green),
                lngRGB(lngCntr + 1) = lngRGB(lngCntr + 1) + lngPorM(lngCntr + 1)         '  then add 1 to next color group
            End If
        End If
    Next lngCntr

End Sub

Private Sub subCurFreq()

    With picSpeed
        .Move ((Me.Width \ Screen.TwipsPerPixelX) - .Width) \ 2, (Me.Height \ Screen.TwipsPerPixelY) - .Height - 21
        hscSpeed.Value = 50
    End With  '  picSpeed

    lngSpeed = hscSpeed.Value / 10

    QueryPerformanceFrequency curFrequency
    curFrequency = curFrequency / 357.9545

End Sub

Private Sub subNewColor()

Dim lngLoop As Long

    For lngLoop = 1 To 3     '  Setup the initial Color Values and direction of change
        lngRGB(lngLoop) = 256 * Rnd
        lngPorM(lngLoop) = Sgn(Rnd - 0.5)   '  A Plus 1 or Minus 1 - lngPorM neans Plus or Minus
    Next lngLoop

End Sub

Private Sub subNewWand()

Dim lngLoop As Long

    For lngLoop = 1 To 4     '  Setup the initial Color Values and direction of change
'       Set initial Direction
        lngDir(lngLoop) = Sgn(Rnd - 0.5)    '  A Plus 1 or Minus 1
'       Set Speed
        lngAngle(lngLoop) = (Rnd * 2) + 1
'       Set initial Location
        lngXY(lngLoop) = Me.ScaleHeight * Rnd
    Next lngLoop

    Me.Cls
    blnWand = False
    mnuMenu.Visible = False

    subNewColor

    If mnuDeskTop.Checked = True Then
        PaintDesktop Me.hdc
    Else
        Cls
    End If

    subWand

End Sub

Private Sub subShowMouse()

'   Show the Mouse Pointer
    tmrMouse.Enabled = False
    MousePointer = 0
    tmrMouse.Enabled = True

End Sub

Private Sub subWand()

Dim lngCntr     As Long
Dim lngColorSeq As Long
Dim lngLoop     As Long
Dim lngStart    As Long

    lngStart = GetTickCount + 60000         '  Run the display for 1 minute
    lngColorSeq = Int(Rnd * 6) + 1

    Do
        With Me
            If Not blnTrail Then
'               Erase existing line
                .ForeColor = vbBlack
            Else
                Select Case lngColorSeq
                Case 1
                    .ForeColor = RGB(lngRGB(1), lngRGB(2), lngRGB(3))
                Case 2
                    .ForeColor = RGB(lngRGB(1), lngRGB(3), lngRGB(2))
                Case 3
                    .ForeColor = RGB(lngRGB(2), lngRGB(1), lngRGB(3))
                Case 4
                    .ForeColor = RGB(lngRGB(2), lngRGB(3), lngRGB(1))
                Case 5
                    .ForeColor = RGB(lngRGB(3), lngRGB(1), lngRGB(2))
                Case 6
                    .ForeColor = RGB(lngRGB(3), lngRGB(2), lngRGB(1))
                End Select   '  lngColorSeq
            End If

            If lngSpeed > 1 Then
                If blnWand Then
'                   Draw over the leading edge
'                   This will keep the color flow consistent
                    MoveToEx .hdc, lngXY(5), lngXY(6), ByVal 0&
                    LineTo .hdc, lngXY(7), lngXY(8)
                Else
                    blnWand = True
                End If

'               This makes the leading edge color opposite to the rest of the display
                Select Case lngColorSeq
                Case 1
                    .ForeColor = RGB(lngRGB(2), lngRGB(3), lngRGB(1))
                Case 2
                    .ForeColor = RGB(lngRGB(2), lngRGB(1), lngRGB(3))
                Case 3
                    .ForeColor = RGB(lngRGB(3), lngRGB(2), lngRGB(1))
                Case 4
                    .ForeColor = RGB(lngRGB(3), lngRGB(1), lngRGB(2))
                Case 5
                    .ForeColor = RGB(lngRGB(1), lngRGB(2), lngRGB(3))
                Case 6
                    .ForeColor = RGB(lngRGB(1), lngRGB(3), lngRGB(2))
                End Select   '  lngColorSeq
            End If

'           Draw new line
            MoveToEx .hdc, lngXY(1), lngXY(2), ByVal 0&
            LineTo .hdc, lngXY(3), lngXY(4)
        End With             '  Me

        For lngLoop = 1 To 4
'           Save the current location
            lngXY(lngLoop + 4) = lngXY(lngLoop)
'           Create the new location
            lngXY(lngLoop) = lngXY(lngLoop) + (lngDir(lngLoop) * lngAngle(lngLoop))
        Next lngLoop

'       Handle bouncing (change directions if one end of the wand is at the edge of the screen)
        For lngCntr = 1 To 4
            If lngXY(lngCntr) < 1 Then
                lngXY(lngCntr) = 1
                lngDir(lngCntr) = -lngDir(lngCntr)
            End If
        Next lngCntr

        For lngCntr = 1 To 3 Step 2
            If lngXY(lngCntr) > Me.ScaleWidth Then
                lngXY(lngCntr) = Me.ScaleWidth - 1
                lngDir(lngCntr) = -lngDir(lngCntr)
            End If

            If lngXY(lngCntr + 1) > Me.ScaleHeight Then
                lngXY(lngCntr + 1) = Me.ScaleHeight
                lngDir(lngCntr + 1) = -lngDir(lngCntr + 1)
            End If
        Next lngCntr

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    To obtain maximum speed, comment out these 3 lines by        ****
'   ****    putting a Tick (') Mark in front of each line.               ****
'   ****                                                                 ****
'   *************************************************************************
        QueryPerformanceCounter curStart    '  Get the start time

        Do While curEnd < curStart + (lngSpeed * curFrequency)
            QueryPerformanceCounter curEnd  '  Get the end time
'   *************************************************************************
'   ****                                                                 ****
'   ****    Do NOT comment out these three lines!                        ****
'   ****                                                                 ****
'   *************************************************************************
            If GetQueueStatus(QS_ALLINPUT) Then
                DoEvents            '  Yield for other processing.
            End If
'   *************************************************************************
'   ****                                                                 ****
'   ****    Comment out this line for maximum speed.                     ****
'   ****                                                                 ****
'   *************************************************************************
        Loop
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Now run the program again and keep watching the colors       ****
'   ****    around the edges.                                            ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
        subChgColors
    Loop Until GetTickCount > lngStart

    subNewWand

End Sub

Private Sub tmrMouse_Timer()

'   Hide the Mouse Pointer
    MousePointer = 99

End Sub

Private Sub tmrSpeed_Timer()

    SetCursorPos 2000, 2000
    picSpeed.Visible = False
    tmrSpeed.Enabled = False

End Sub

':)Code Fixer V4.0.31 (Saturday, 10 June 2006 01:56:49) 82 + 405 = 487 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|033322232223332222222222222222222|1113222|2221222|333332222233|111111111111|1122222222222|333333|
