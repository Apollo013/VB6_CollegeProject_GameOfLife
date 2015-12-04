VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Game_Of_Life 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Of Life - Paul Millar ..."
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   846
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frme_Counter 
      Caption         =   "Counter ..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1290
      Left            =   8475
      TabIndex        =   13
      Top             =   3900
      Width           =   4050
      Begin VB.TextBox txt_LiveCellCount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2200
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   1545
      End
      Begin VB.TextBox txt_StepCounter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2200
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label6 
         Caption         =   "Live Cells:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   16
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Current Step ..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   270
         TabIndex        =   15
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.CommandButton btn_Save 
      Height          =   750
      Left            =   10125
      Picture         =   "FRM_Game_Of_Life.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save Game ..."
      Top             =   90
      Width           =   750
   End
   Begin VB.CommandButton btn_Open 
      Height          =   750
      Left            =   9300
      Picture         =   "FRM_Game_Of_Life.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Open Game ..."
      Top             =   75
      Width           =   750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8460
      Top             =   7755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frme_Step 
      Caption         =   "Step ..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   8475
      MouseIcon       =   "FRM_Game_Of_Life.frx":6904
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   945
      Width           =   4050
      Begin MSComctlLib.Slider sld_Step 
         Height          =   420
         Left            =   90
         TabIndex        =   11
         Top             =   285
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
         _Version        =   393216
         LargeChange     =   50
         Max             =   500
         TickFrequency   =   50
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3060
         TabIndex        =   12
         Top             =   810
         Width           =   900
      End
   End
   Begin VB.CommandButton btn_New 
      Height          =   750
      Left            =   8475
      Picture         =   "FRM_Game_Of_Life.frx":6A56
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "New Game ..."
      Top             =   90
      Width           =   750
   End
   Begin VB.CommandButton btn_Stop 
      Height          =   750
      Left            =   11775
      Picture         =   "FRM_Game_Of_Life.frx":9ED8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Stop Game ..."
      Top             =   90
      Width           =   750
   End
   Begin VB.CommandButton btn_Start 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   10950
      Picture         =   "FRM_Game_Of_Life.frx":D35A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Start Game ..."
      Top             =   90
      Width           =   750
   End
   Begin VB.PictureBox pct_GameGrid 
      Height          =   8325
      Left            =   45
      ScaleHeight     =   551
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   0
      Top             =   60
      Width           =   8325
   End
   Begin VB.Frame pnl_ColourPallete 
      Caption         =   "Colour Palette ..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   8475
      MouseIcon       =   "FRM_Game_Of_Life.frx":107DC
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2160
      Width           =   4050
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Line Colour"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   2835
         TabIndex        =   7
         Top             =   1000
         Width           =   885
      End
      Begin VB.Shape shp_GridLineColour 
         BackStyle       =   1  'Opaque
         Height          =   520
         Left            =   2800
         Top             =   380
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Live Cell Colour"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1605
         TabIndex        =   6
         Top             =   1000
         Width           =   840
      End
      Begin VB.Shape shp_LiveCellColour 
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   1530
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Background Colour"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   210
         TabIndex        =   5
         Top             =   1000
         Width           =   1080
      End
      Begin VB.Shape shp_BackgroundColour 
         BackStyle       =   1  'Opaque
         Height          =   520
         Left            =   280
         Top             =   380
         Width           =   900
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_File_New 
         Caption         =   "&New"
      End
      Begin VB.Menu mnu_File_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnu_File_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Game 
      Caption         =   "&Game"
      Begin VB.Menu mnu_Game_Start 
         Caption         =   "Sta&rt"
      End
      Begin VB.Menu mnu_Game_Stop 
         Caption         =   "S&top"
      End
   End
End
Attribute VB_Name = "FRM_Game_Of_Life"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oGof    As CLS_Game_Of_Life

Private Sub Form_Activate()
'--------------------------------------------------------------------------------------------
' DATE:         07/03/2011
' DESCRIPTION:  Initialise Game.
'--------------------------------------------------------------------------------------------
    'Create a new Game Of Life Object.
    Set oGof = New CLS_Game_Of_Life
    
    'Assign main picture control and counter controls to allow event and property handling.
    oGof.GameGrid = pct_GameGrid
    oGof.CounterControl = txt_StepCounter
    oGof.CellCounterControl = txt_LiveCellCount
    
    'Common Dialog Control Is Needed When We Save Or Open A File.
    oGof.GameDialog = CommonDialog1
    
    'Start A New Game.
    oGof.New_Game
    
    'Initialise Slider Control
    sld_Step.Min = 0
    sld_Step.Max = 500
    sld_Step.LargeChange = 50
    sld_Step.value = 0
    
    'Update controls on right panel
    Form_Update
End Sub

Private Sub Form_Deactivate()
'--------------------------------------------------------------------------------------------
' DATE:         07/03/2011
' DESCRIPTION:  De-reference Game Of Life Object.
'--------------------------------------------------------------------------------------------
    Set oGof = Nothing
End Sub

Private Sub Form_Update()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Updates Visuals On Some Form Controls.
'--------------------------------------------------------------------------------------------
    'Title bar
    Me.Caption = "Game Of Life - " & oGof.FileName
    'These are the color panels.
    shp_BackgroundColour.BackColor = oGof.BackgroundColour
    shp_LiveCellColour.BackColor = oGof.LiveCellColour
    shp_GridLineColour.BackColor = oGof.GridLineColour
End Sub

Private Sub btn_New_Click()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Initialises A New Game.
'--------------------------------------------------------------------------------------------
    oGof.New_Game
    Form_Update
End Sub

Private Sub btn_Open_Click()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Opens A Pre-Saved Game.
'--------------------------------------------------------------------------------------------
    oGof.Open_Game
    Form_Update
End Sub

Private Sub btn_Start_Click()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Starts A New Game.
'               Passes The Value Of The Slider Control. This Determines The 'Step' Value.
'--------------------------------------------------------------------------------------------
    oGof.Start_Game sld_Step.value
End Sub

Private Sub btn_Stop_Click()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Stops Currently Running Game.
'--------------------------------------------------------------------------------------------
    oGof.Stop_Game
End Sub

Private Sub btn_Save_Click()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Saves Current Game.
'--------------------------------------------------------------------------------------------
    oGof.Save_Game
    Form_Update
End Sub

'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Menu Items.
'--------------------------------------------------------------------------------------------
Private Sub mnu_File_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_File_New_Click()
    btn_New_Click
End Sub

Private Sub mnu_File_Open_Click()
    btn_Open_Click
End Sub

Private Sub mnu_File_Save_Click()
    btn_Save_Click
End Sub

Private Sub mnu_Game_Start_Click()
    btn_Start_Click
End Sub

Private Sub mnu_Game_Stop_Click()
    btn_Stop_Click
End Sub

Private Sub pnl_ColourPallete_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Determines If User Has Clicked Inside One Of The 3 Colour Palletes.
'               If So, We Pass An Enumerated Value To the Class Method 'Change_Colour_Scheme'
'               This in Turn Calls The Colour Common Dialog.
'--------------------------------------------------------------------------------------------
    If (X >= shp_BackgroundColour.Left And X <= (shp_BackgroundColour.Left + shp_BackgroundColour.Width)) _
        And (y >= shp_BackgroundColour.Top And y <= (shp_BackgroundColour.Top + shp_BackgroundColour.Height)) Then
        
        oGof.Change_Colour_Scheme (Background)
    
    ElseIf (X >= shp_LiveCellColour.Left And X <= (shp_LiveCellColour.Left + shp_LiveCellColour.Width)) _
        And (y >= shp_LiveCellColour.Top And y <= (shp_LiveCellColour.Top + shp_LiveCellColour.Height)) Then
        
        oGof.Change_Colour_Scheme (Cell)
            
    ElseIf (X >= shp_GridLineColour.Left And X <= (shp_GridLineColour.Left + shp_GridLineColour.Width)) _
        And (y >= shp_GridLineColour.Top And y <= (shp_GridLineColour.Top + shp_GridLineColour.Height)) Then
        
        oGof.Change_Colour_Scheme (GridLine)
    End If
    
    Form_Update
End Sub

Private Sub sld_Step_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Sets Focus On The Grid, Simply Because I Didn't Like The Black Square Line
'               Left Around The Slider Control After Clicking On It.
'--------------------------------------------------------------------------------------------
    pct_GameGrid.SetFocus
End Sub

Private Sub sld_Step_Scroll()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Displays The Value Of The Slider Control.
'--------------------------------------------------------------------------------------------
    Label4.Caption = sld_Step.value
End Sub

Private Sub txt_LiveCellCount_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Prohibit User From Entering Any Data.
'--------------------------------------------------------------------------------------------
    Beep
    KeyAscii = 0
End Sub

Private Sub txt_StepCounter_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Prohibit User From Entering Any Data.
'--------------------------------------------------------------------------------------------
    Beep
    KeyAscii = 0
End Sub
