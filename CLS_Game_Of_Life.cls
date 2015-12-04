VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLS_Game_Of_Life"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ACTIVEX CONTROL REFERENCES
Private WithEvents txt_PictureControl       As VB.PictureBox
Attribute txt_PictureControl.VB_VarHelpID = -1
Private WithEvents txt_StepCounterControl   As VB.TextBox
Attribute txt_StepCounterControl.VB_VarHelpID = -1
Private WithEvents txt_CellCounterControl   As VB.TextBox
Attribute txt_CellCounterControl.VB_VarHelpID = -1
Private WithEvents CommonDialogControl      As MSComDlg.CommonDialog
Attribute CommonDialogControl.VB_VarHelpID = -1

' CLASS REFERENCES
Private oFileAccess As New CLS_File_Access

' CONSTANTS
Const DefaultFileName = "[New Game]"

' GAME PROPERTY DECLARATIONS
Private iGameStatus         As enmGameStatus
Private sFileName           As String
Private sGameFolder         As String

' COUNTER PROPERTY DECLARATIONS
Private iXCellCount         As Integer
Private iYCellCount         As Integer
Private lCurrentStepCount   As Long
Private lLiveCellCount      As Long
Private iArrayXCounter      As Integer
Private iArrayYCounter      As Integer

' ARRAY DECLARATIONS
Private arrLiveCells()      As Boolean  'Tracks If Cell Is Currently Dead Or Alive.
Private arrCellTally()      As Integer  'Tallys Neighbouring Live Cell Count For Individual Cells.

' GRID PROPERTIES
Private lBackgroundColour   As Long
Private lGridLineColour     As Long
Private lLiveCellColour     As Long
Private iCellSize           As Integer
Private iGridLineWidth      As Integer
Private iFullCellSize       As Integer

' ENUMERATORS
Public Enum enmGameStatus
    Unchanged
    Changed
    Running
    Stopped
    Saved
    Opened
End Enum

Public Enum enmGridAttribute
    Background
    GridLine
    Cell
End Enum

Public Enum enmGridDrawOption
    WithoutCells
    WithCurrentCells
End Enum

'----------------------------------------------------------------
' GAME PROPERTIES
'----------------------------------------------------------------
Public Property Get FileName() As String
    FileName = sFileName
End Property
Public Property Let FileName(value As String)
    sFileName = value
End Property

'----------------------------------------------------------------
' ACTIVEX PROPERTIES
'----------------------------------------------------------------
Public Property Get GameGrid() As VB.PictureBox
    Set GameGrid = txt_PictureControl
End Property
Public Property Let GameGrid(pGrid As VB.PictureBox)
    Set txt_PictureControl = pGrid
End Property

Public Property Let GameDialog(value As MSComDlg.CommonDialog)
    Set CommonDialogControl = value
End Property

Public Property Let CounterControl(value As VB.TextBox)
    Set txt_StepCounterControl = value
End Property

Public Property Let CellCounterControl(value As VB.TextBox)
    Set txt_CellCounterControl = value
End Property

'----------------------------------------------------------------
' GRID PROPERTIES
'----------------------------------------------------------------
Public Property Get BackgroundColour() As Long
    BackgroundColour = lBackgroundColour
End Property
Public Property Let BackgroundColour(value As Long)
    lBackgroundColour = value
End Property

Public Property Get GridLineColour() As Long
    GridLineColour = lGridLineColour
End Property
Public Property Let GridLineColour(value As Long)
    lGridLineColour = value
End Property

Public Property Get LiveCellColour() As Long
    LiveCellColour = lLiveCellColour
End Property
Public Property Let LiveCellColour(value As Long)
    lLiveCellColour = value
End Property

'----------------------------------------------------------------
' METHODS
'----------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------
' DATE:         07/03/2011
' DESCRIPTION:  De-Reference Any Objects.
'--------------------------------------------------------------------------------------------
    Set oFileAccess = Nothing
End Sub

Public Sub New_Game()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Initialises A New Game.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    'Check If The Current Game Has Any Changes That Need Saving.
    Check_Save
    
    'Re-Initialise Variables.
    Set_Defaults
    
    'Re-Dimension Arrays
    ReDim arrLiveCells(iXCellCount + 1, iYCellCount + 1)
    ReDim arrCellTally(iXCellCount + 1, iYCellCount + 1)
    
    'Draw The Grid A-Fresh (Do Not Draw Any Live Cells).
    Draw_Grid
    
    'Display Current Step Count And Number Of Live Cells To The User
    Update_ActiveX_Counter_Controls
    
    'Reset All Array Values To False.
    For iArrayXCounter = 1 To iXCellCount
        For iArrayYCounter = 1 To iYCellCount
            arrLiveCells(iArrayXCounter, iArrayYCounter) = False
        Next iArrayYCounter
    Next iArrayXCounter
End Sub

Private Sub Set_Defaults()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Re-Initialise Variables.
'--------------------------------------------------------------------------------------------
    sGameFolder = App.Path & "\GameData"
    iGameStatus = Unchanged
    sFileName = DefaultFileName
    
    iXCellCount = 50
    iYCellCount = 50
    lCurrentStepCount = 0
    lLiveCellCount = 0
    
    If Me.BackgroundColour = 0 Then Me.BackgroundColour = vbWhite
    If Me.GridLineColour = 0 Then Me.GridLineColour = &H808080
    If Me.LiveCellColour = 0 Then Me.LiveCellColour = &HFF0000
    
    iCellSize = 10
    iGridLineWidth = 1
    iFullCellSize = iCellSize + iGridLineWidth
End Sub

Public Sub Open_Game()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Opens A Pre-Existing Game.
'               All Games Are Stored In A Sub Folder Of The Current Directory.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
        
    'Check If The Current Game Has Any Changes That Need Saving.
    Check_Save
    
    'Check That Game Folder Exists, If Not Then Create It.
    oFileAccess.Check_Folder sGameFolder, ForceCreate
    
    'Open File Dialog
    CommonDialogControl.DialogTitle = "Open File ..."
    CommonDialogControl.FileName = ""
    CommonDialogControl.Filter = "*.GOF (Game Of Life Files)|*.GOF"
    CommonDialogControl.InitDir = sGameFolder
    CommonDialogControl.ShowOpen
    sFileName = CommonDialogControl.FileName

    'Check That We Have An Appropriate Name.
    If sFileName <> "" And sFileName <> DefaultFileName Then
        'Open Data File For INPUT.
        oFileAccess.Open_File For_Input, sFileName
        
        'Populate Array.
        For iArrayXCounter = 1 To iXCellCount
            For iArrayYCounter = 1 To iYCellCount
                Input #oFileAccess.FileHandle, arrLiveCells(iArrayXCounter, iArrayYCounter)
            Next iArrayYCounter
        Next iArrayXCounter
        
        'Close Data File.
        oFileAccess.Close_File
        
        'Set Game Status.
        iGameStatus = Opened
        
        'Redraw The Grid With Current Values In Array.
        Draw_Grid WithCurrentCells
    End If
End Sub

Private Sub Check_Save()
'08
    If iGameStatus = Changed Then
        If MsgBox("This Game Has Been Changed." & vbCrLf & "Do You Want To Save Your Changes ?", vbYesNo + vbQuestion, "Save Game ?") = vbYes Then
            Me.Save_Game
        End If
    End If
End Sub

Public Sub Save_Game()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Saves Current Game.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    'Check That Game Folder Exists, If Not Then Create It.
    oFileAccess.Check_Folder sGameFolder, ForceCreate
    
    'Open File Dialog
    If sFileName = "" Or sFileName = DefaultFileName Then
        CommonDialogControl.DialogTitle = "Save As ..."
        CommonDialogControl.FileName = ""
        CommonDialogControl.Filter = "*.GOF (Game Of Life Files)|*.GOF"
        CommonDialogControl.InitDir = sGameFolder
        CommonDialogControl.ShowSave
        sFileName = CommonDialogControl.FileName
    End If

    'Check That We Have An Appropriate Name.
    If sFileName = "" Or sFileName = DefaultFileName Then
        MsgBox "A Valid File Name Must Be Provided." & vbCrLf & "Please Try Saving Again.", vbInformation + vbOKOnly, "Invalid File Name"
        Exit Sub
    Else
    
        'Make Sure We Have The The ".GOF" Extension.
        If Right(Trim(sFileName), 4) <> ".GOF" Then
            sFileName = sFileName & ".GOF"
        End If
        
        'Open File For OUPUT.
        oFileAccess.Open_File For_Output, sFileName
        
        'Write Array Values.
        For iArrayXCounter = 1 To iXCellCount
            For iArrayYCounter = 1 To iYCellCount
                Write #oFileAccess.FileHandle, arrLiveCells(iArrayXCounter, iArrayYCounter)
            Next iArrayYCounter
        Next iArrayXCounter
        
        'Close File.
        oFileAccess.Close_File
        
        'Change Game Status To Saved.
        iGameStatus = Saved
    End If
End Sub

Public Sub Start_Game(Optional ByVal pStep As Long)
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Saves Current Game.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    'Additional Check
    If lLiveCellCount = 0 Then Exit Sub
    
    Dim lStep As Long
    
    'Essentially Check If lCurrentStepCount Exceeds System Limit.
    On Error GoTo Start_Game_ERROR
    
    'Set Game Status To Running.
    iGameStatus = Running
    
    If IsMissing(pStep) Or pStep = 0 Then
        'No Step Value Provided So Iterate Infinately(?)
        Do While iGameStatus = Running
            lCurrentStepCount = lCurrentStepCount + 1
            Count_Live_Cells
            Populate_Next_Generation
            'DoEvents - Allows The User To Stop The Game Manually.
            DoEvents
        Loop
    Else
        'Step Value Provided So Iterate Step(n)
        For lStep = 1 To pStep
            lCurrentStepCount = lCurrentStepCount + 1
            Count_Live_Cells
            Populate_Next_Generation
            'DoEvents - Allows The User To Stop The Game Manually.
            DoEvents
            'Check If User Has Manually Stopped The Game
            If iGameStatus = Stopped Then Exit For
        Next lStep
        iGameStatus = Stopped
    End If
    
    'Change Game Status To Changed.
    iGameStatus = Changed
    
Start_Game_OUT:
    Exit Sub
    
Start_Game_ERROR:
    If Err.Number Then
        MsgBox "Error Number: " & Err.Number & vbCrLf & _
                Err.Description, vbInformation + vbOKOnly, "Error"
        iGameStatus = Stopped
        Resume Start_Game_OUT
    End If
End Sub

Public Sub Stop_Game()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Stops Current Game.
'               The Method 'Start_Game' Will Pick Up On The Value For Game Status And Stop.
'--------------------------------------------------------------------------------------------
    iGameStatus = Stopped
End Sub

Private Sub Draw_Grid(Optional ByVal GridOption As enmGridDrawOption = 0)
'--------------------------------------------------------------------------------------------
' DATE:         07/03/2011
' DESCRIPTION:  Draws The Grid.
'               The Arguement 'GridOption' Specifies Whether Or Not We Draw Any Live Cells.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    Dim iCounter        As Integer
    
    'Start With A Fresh Canvas.
    Me.GameGrid.Cls
    
    'Set Backgound Colour.
    Me.GameGrid.BackColor = Me.BackgroundColour
    
    'Set Picture Box Height And Width Properties
    Me.GameGrid.Height = (iYCellCount * iFullCellSize) + 5
    Me.GameGrid.Width = (iXCellCount * iFullCellSize) + 5
        
    'Draw Horizontal Grid Lines
    For iCounter = 0 To iYCellCount
        Me.GameGrid.Line (0, (iCounter * iFullCellSize))-((Me.GameGrid.Width), (iCounter * iFullCellSize)), Me.GridLineColour
    Next iCounter
        
    'Draw Vertical Grid Lines
    For iCounter = 0 To iXCellCount
        Me.GameGrid.Line ((iCounter * iFullCellSize), 0)-((iCounter * iFullCellSize), (Me.GameGrid.Height)), Me.GridLineColour
    Next iCounter
    
    If GridOption = WithCurrentCells Then
        'Redraw Any Live Cells.
        For iArrayXCounter = 1 To iXCellCount
            For iArrayYCounter = 1 To iYCellCount
                If arrLiveCells(iArrayXCounter, iArrayYCounter) = True Then
                    Fill_Cell iArrayXCounter, iArrayYCounter, Me.LiveCellColour
                End If
            Next iArrayYCounter
        Next iArrayXCounter
    End If
End Sub

Private Sub txt_PictureControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'--------------------------------------------------------------------------------------------
' DATE:         07/03/2011
' DESCRIPTION:  Traps User Click On Picture Control And Colours In Individual Cells
'               Depending On X,Y Coordinates.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    Dim iXCellNumber    As Integer
    Dim iYCellNumber    As Integer
    Dim lTempColour     As Long
    
    'Trap User Click Position Horizontally.
    iXCellNumber = Int(X / iFullCellSize) + 1
    If iXCellNumber > iXCellCount Then iXCellNumber = iXCellCount
    If iXCellNumber < 1 Then iXCellNumber = 1
    
    'Trap User Click Position Vertically.
    iYCellNumber = Int(y / iFullCellSize) + 1
    If iYCellNumber > iYCellCount Then iYCellNumber = iYCellCount
    If iYCellNumber < 1 Then iYCellNumber = 1
    
    'Store Cell Status
    arrLiveCells(iXCellNumber, iYCellNumber) = Not (arrLiveCells(iXCellNumber, iYCellNumber))
    
    'Determine Cell Fill Colour To Use (Live Cell Colour Our Background Colour).
    If arrLiveCells(iXCellNumber, iYCellNumber) Then
        lTempColour = Me.LiveCellColour
    Else
        lTempColour = Me.BackgroundColour
    End If
    
    'Fill In Selected Grid Cell.
    Fill_Cell iXCellNumber, iYCellNumber, lTempColour
    
    'Change Game Status.
    iGameStatus = Changed
End Sub

Public Sub Change_Colour_Scheme(ByVal pGridAttribute As enmGridAttribute)
'--------------------------------------------------------------------------------------------
' DATE:         09/03/2011
' DESCRIPTION:  Calls The Colour Dialog And Sets Colour Values For either Background,
'               Grid Line or Live Cell.
'--------------------------------------------------------------------------------------------
    'Check All Is In Order Before Proceeding.
    If Not Check_Game_Status Then Exit Sub
    
    Dim Colour As Long
    
    If pGridAttribute = Background Then
        CommonDialogControl.DialogTitle = "Select Background Colour"
    ElseIf pGridAttribute = GridLine Then
        CommonDialogControl.DialogTitle = "Select Grid Line Colour"
    Else
        CommonDialogControl.DialogTitle = "Select Live Cell Colour"
    End If
    
    'CommonDialogControl.DialogTitle = "Choose Colour"
    CommonDialogControl.ShowColor
    Colour = CommonDialogControl.Color
    
    If Colour <> 0 Then
        If pGridAttribute = Background Then
            Me.BackgroundColour = Colour
        ElseIf pGridAttribute = GridLine Then
            Me.GridLineColour = Colour
        Else
            Me.LiveCellColour = Colour
        End If
    End If
    
    'Redraw The Grid With Any Live Cells.
    Draw_Grid WithCurrentCells
End Sub

Private Function Check_Game_Status() As Boolean
'--------------------------------------------------------------------------------------------
' DATE:         09/03/2011
' DESCRIPTION:  Checks That The Game Grid Property Was Initialised And That The Game Is Not
'               Currently Running. Returns FALSE If Either Of These Is true.
'--------------------------------------------------------------------------------------------
    Check_Game_Status = True
    
    If Me.GameGrid Is Nothing Or iGameStatus = Running Then
        Check_Game_Status = False
    End If
End Function

Private Sub Update_ActiveX_Counter_Controls()
'--------------------------------------------------------------------------------------------
' DATE:         09/03/2011
' DESCRIPTION:  Display Current Step Count And Number Of Live Cells To The User.
'--------------------------------------------------------------------------------------------
    txt_StepCounterControl.Text = lCurrentStepCount
    txt_CellCounterControl.Text = lLiveCellCount
End Sub

Private Sub Count_Live_Cells()
'--------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Counts How Many Live Cells Directly Surround Each Cell And The Writes That
'               Result To The Array 'arrCellTally'.
'--------------------------------------------------------------------------------------------
    For iArrayXCounter = 1 To iXCellCount
        For iArrayYCounter = 1 To iYCellCount
            arrCellTally(iArrayXCounter, iArrayYCounter) = arrLiveCells(iArrayXCounter - 1, iArrayYCounter - 1) + _
                                   arrLiveCells(iArrayXCounter - 1, iArrayYCounter) + _
                                   arrLiveCells(iArrayXCounter - 1, iArrayYCounter + 1) + _
                                   arrLiveCells(iArrayXCounter, iArrayYCounter - 1) + _
                                   arrLiveCells(iArrayXCounter, iArrayYCounter + 1) + _
                                   arrLiveCells(iArrayXCounter + 1, iArrayYCounter - 1) + _
                                   arrLiveCells(iArrayXCounter + 1, iArrayYCounter) + _
                                   arrLiveCells(iArrayXCounter + 1, iArrayYCounter + 1)
            arrCellTally(iArrayXCounter, iArrayYCounter) = -arrCellTally(iArrayXCounter, iArrayYCounter)
        Next iArrayYCounter
    Next iArrayXCounter
End Sub

Private Sub Populate_Next_Generation()
'----------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Works Out The Next Generation Based On The Results In The Array 'arrCellTally'.
'               Calls The Method 'Fill_Cell' To Colour Each Cell Appropriately.
'----------------------------------------------------------------------------------------------
    Dim bGameChanged As Boolean
    
    bGameChanged = False
    
    For iArrayXCounter = 1 To iXCellCount
        For iArrayYCounter = 1 To iYCellCount
            'If Cell Has 3 Surrounding Live Cells.
            If arrCellTally(iArrayXCounter, iArrayYCounter) = 3 Then
                'Only Give Life To A Cell If It Is Currently Dead.
                If Not arrLiveCells(iArrayXCounter, iArrayYCounter) Then
                    arrLiveCells(iArrayXCounter, iArrayYCounter) = True
                    Fill_Cell iArrayXCounter, iArrayYCounter, Me.LiveCellColour
                    'Game Has Changed
                    bGameChanged = True
                End If
            'If Cell Is Under or Over Populated Then Kill It.
            ElseIf arrCellTally(iArrayXCounter, iArrayYCounter) <= 1 Or arrCellTally(iArrayXCounter, iArrayYCounter) >= 4 Then
                'Only Kill It If It Is Currently Alive (?).
                If arrLiveCells(iArrayXCounter, iArrayYCounter) Then
                    arrLiveCells(iArrayXCounter, iArrayYCounter) = False
                    Fill_Cell iArrayXCounter, iArrayYCounter, Me.BackgroundColour
                    'Game Has Changed
                    bGameChanged = True
                End If
            End If
        Next iArrayYCounter
    Next iArrayXCounter
    
    'If Game Has Not Changed Then Inform User And Stop Game.
    If Not bGameChanged Then
        MsgBox "No More Generations Can Be Populated." & vbCrLf & "The Game Has Ended.", vbInformation + vbOKOnly, "Game Of Life"
        iGameStatus = Stopped
    End If
End Sub

Private Sub Fill_Cell(X As Integer, y As Integer, Colour As Long)
'----------------------------------------------------------------------------------------------
' DATE:         08/03/2011
' DESCRIPTION:  Fills Individual Cells Based On The Coordinate Arguments And The Colour Arg.
'----------------------------------------------------------------------------------------------
    'Fill In Selected Grid Cell.
    Me.GameGrid.Line ((X - 1) * (iFullCellSize) + iGridLineWidth, _
                     (y - 1) * (iFullCellSize) + iGridLineWidth)- _
                     Step(iCellSize - 1, iCellSize - 1), Colour, BF
                     
    'Tally The Number Of Live Cells There Currently Are.
    If Colour = BackgroundColour Then
        lLiveCellCount = lLiveCellCount - 1
        If lLiveCellCount < 0 Then lLiveCellCount = 0
    Else
        lLiveCellCount = lLiveCellCount + 1
    End If
    
    'Display Current Step Count And Number Of Live Cells To The User
    Update_ActiveX_Counter_Controls
End Sub



