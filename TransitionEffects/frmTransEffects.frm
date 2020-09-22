VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Transition Effects"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   Icon            =   "frmTransEffects.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Stretch Wipe In Hide V "
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   26
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Stretch Wipe In Push H"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Bar2 Right"
      Height          =   495
      Index           =   3
      Left            =   6240
      TabIndex        =   24
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Bar2 Left"
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   23
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Bar2 Down"
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   22
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Bar2  Up"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   21
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Bars Move Horizontal"
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Bars Move Vertical"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Bar Draw Horizontal"
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Bar Draw Vertical"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Wipe Out vertical"
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Wipe Out vertical"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Wipe In Horizontal"
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Wipe In Veritcal"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Right"
      Height          =   495
      Index           =   3
      Left            =   6240
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Left"
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Down"
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Up"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stretching Move"
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Slide -Down"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Lines Horizontal"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stretching Push"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Slide -Up"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   6840
      Picture         =   "frmTransEffects.frx":0442
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Lines Vertical"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   120
      Picture         =   "frmTransEffects.frx":53CF
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   120
      Width           =   3885
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   6720
      Picture         =   "frmTransEffects.frx":A35C
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Transition Effects By Mohammed Ali Sohrabi ,ali6236@yahoo.com
'   Cool Transition for your programs!
'   See Notes on the module
Public StopProgram As Boolean
Private Sub Command1_Click(Index As Integer)
'Random Lines
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical - Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Enable
'   Step             : Disable
'**********************************
'   Notes:
'   RefreshRate : number of lines in each refresh
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        RandomLines Picture1, Picture2, VerticalSide, 0
    Else
        'the speed is 1, but it is slow, we use RefreshRate for faster result...
        RandomLines Picture1, Picture2, HorizontalSide, 2
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command10_Click(Index As Integer)
'**** when i run this section my computer restarts!
'**** please enable and run, and say me about your computer....

'Stretching
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : Yes
'   Sides            : Vertical - Horizontal
'**********************************
'   Push Modes       : Enable (Push,Hide)
'   Refresh Rate     : Enable
'   Step             : Enable
'**********************************
'   Notes:
'   Stretch is a slow effect,
'   Just use for small pictures, with large steps
'   and use push mode just when you need,
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        Stretching_Wipe_In Picture1, Picture3, Picture2, HorizontalSide, 5, 0, Pushing
    Else
        Stretching_Wipe_In Picture1, Picture3, Picture2, VerticalSide, 5, 0, Hiding
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command2_Click(Index As Integer)
'Slide
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : Yes
'   Sides            : All Sides (Up and Down are completed)
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes: Just use Up and Down,
'          I will complete other sides as soon as possible!
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        Slide Picture1, Picture3, Picture2, aUp, 3
    Else
        Slide Picture1, Picture3, Picture2, aDown, 3
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command3_Click(Index As Integer)
'Stretching
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : Yes
'   Sides            : Left - Right
'**********************************
'   Push Modes       : Enable (Push,Move)
'   Refresh Rate     : Enable
'   Step             : Enable
'**********************************
'   Notes:
'   Stretch is a slow effect,
'   Just use for small pictures, with large steps
'   and use push mode just when you need,
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        Stretching Picture1, Picture3, Picture2, sRight, 5, 0, Pushing
    Else
        Stretching Picture1, Picture3, Picture2, sLeft, 5, 0, Moving
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command4_Click(Index As Integer)
'Wipe
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : All (Left,Right,Up,Down)
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
If Not IsReady Then Exit Sub
    lngSpeed = 1
    Wipe Picture1, Picture2, 2 ^ Index, 3 '!!!! i'm using ^ !!! it's better to use select case
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command5_Click(Index As Integer)
'Wipe In
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical and Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes:
'   This is like two normal wipe.

If Not IsReady Then Exit Sub
    lngSpeed = 1
    Wipe_In Picture1, Picture2, Index + 1, 3
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command6_Click(Index As Integer)
'Wipe Out
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical and Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes:
'   like wipe in....

If Not IsReady Then Exit Sub
    lngSpeed = 1
    Wipe_Out Picture1, Picture2, Index + 1, 3
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command7_Click(Index As Integer)
'Bar Draw
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical and Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes:

If Not IsReady Then Exit Sub
    lngSpeed = 1
    Static way As Boolean
    Bars_Draw Picture1, Picture2, Index + 1, 5, 15
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command8_Click(Index As Integer)
'Bar Move
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical and Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes:
'   like wipe in....

If Not IsReady Then Exit Sub
    lngSpeed = 1
    Static way As Boolean
    Bars_Move Picture1, Picture2, Index + 1, 4, 10
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command9_Click(Index As Integer)
'Wipe
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : All (Left,Right,Up,Down)
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
If Not IsReady Then Exit Sub
    lngSpeed = 50
    Bars_OneSide Picture1, Picture2, 2 ^ Index, 1, 20
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Form_Load()
    StopProgram = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnRunning = False
    StopProgram = True
    Unload Me
End Sub

