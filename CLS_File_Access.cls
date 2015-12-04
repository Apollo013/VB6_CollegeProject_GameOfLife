VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLS_File_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------
' DATE:             05/02/2011
' DESCRIPTION:      Data Access Class
'----------------------------------------------------------------------------------------------------

' API Function For Creating A New Folder (If It Does Not Already Exist)
Private Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

'Indicates If The File Is Opened Or Not.
Private bFileStatus         As enmFileStatus
'Sets Or Returns The Current FreeFile Number.
Private iFileHandle         As Integer
Private sFileName           As String
Private iRecordLength       As Integer

'Different File Access Methods For Data Files.
Public Enum enmFileAccess
    For_Random = 1
    For_Output = 2
    For_Input = 3
End Enum

'File Open Status For Data Files.
Public Enum enmFileStatus
    FileClosed = False
    FileOpened = True
End Enum

Public Enum enmFolderCreation
    DontCreate
    ForceCreate
End Enum

' *** FILE STATUS ***
Public Property Get FileStatus() As enmFileStatus
    FileStatus = bFileStatus
End Property
Public Property Let FileStatus(Status As enmFileStatus)
    bFileStatus = Status
End Property

' *** FILE HANDLE ***
Public Property Get FileHandle() As Integer
    FileHandle = iFileHandle
End Property
Public Property Let FileHandle(Handle As Integer)
    iFileHandle = Handle
End Property

' *** FILE NAME ***
Public Property Get FileName() As String
    FileName = sFileName
End Property
Public Property Let FileName(pFileName As String)
    sFileName = pFileName
End Property

' *** RECORD LENGTH ***
Public Property Get RecordLength() As Integer
    RecordLength = iRecordLength
End Property
Public Property Let RecordLength(value As Integer)
    iRecordLength = value
End Property

' *** RECORD COUNT ***
Public Property Get Count() As Long
    Count = LOF(iFileHandle) / Me.RecordLength
End Property

Private Sub Class_Initialize()
' DO NOT DELETE THIS
End Sub

Private Sub Class_Terminate()
'-------------------------------------------------------------------------------------------------
' DATE:             05/02/2011
' DESCRIPTION:      Closes The Data File.
'-------------------------------------------------------------------------------------------------
    Me.Close_File
End Sub

Public Sub Open_File(ByVal pAccessMethod As enmFileAccess, Optional ByVal pFileName As String, Optional ByVal pRecordLength As Integer)
'-------------------------------------------------------------------------------------------------
' DATE:             05/02/2011
' DESCRIPTION:      Opens The Data File.
'                   A Default Name Is Assigned If The 'FileName' Argument Is Empty Or Unspecified.
' HISTORY:          07/03/2011: Added An Additional Arguement To Parameter List 'pRecordLength' Is
'                               An Integer Value That Specifies The Size Of A 'Type' Structure.
'                               This Must Now Be Provided If We Want To Open A File For Random
'                               Access.
'-------------------------------------------------------------------------------------------------
    On Error GoTo OpenFile_ERR
        
    If pFileName = "" Or IsMissing(pFileName) Then
        Me.FileName = App.Path & "\DataFile.dat"
    Else
        Me.FileName = pFileName
    End If
    
    ' ADDED: 07/03/2011
    If pRecordLength > 0 Then
        Me.RecordLength = pRecordLength
    End If
    
    Me.FileHandle = FreeFile
    
    If pAccessMethod = For_Output Then
        Open Me.FileName For Output As #Me.FileHandle
    ElseIf pAccessMethod = For_Random Then
        ' ALTERED: 07/03/2011
        If Me.RecordLength = 0 Then
            MsgBox "You Must Provide A Record Length If You Want To Open The File For Random Access.", vbOKOnly + vbInformation, "File Open Error"
            GoTo OpenFile_OUT
        Else
            Open Me.FileName For Random As #Me.FileHandle Len = Me.RecordLength
        End If
    ElseIf pAccessMethod = For_Input Then
        Open Me.FileName For Input As #Me.FileHandle
    End If
    
    Me.FileStatus = FileOpened
        
'Just exit
OpenFile_OUT:
    Exit Sub
    
'Can't open file - build error message.
OpenFile_ERR:
    If Err.Number Then
        Select Case Err
            Case 53
                MsgBox "Error 53: File cannot be located.", vbExclamation + vbOKOnly, "File Open Error"
            Case 68
                MsgBox "Error 68: Drive not available.", vbExclamation + vbOKOnly, "File Open Error"
            Case 76
                MsgBox "Error 76: Path not available.", vbExclamation + vbOKOnly, "File Open Error"
            Case Else
                MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, "File Open Error"
        End Select
        'Make sure the file is closed
        Me.Close_File
        Resume OpenFile_OUT
    End If
End Sub

Public Sub Close_File()
'---------------------------------------------------------------------------------------------------------------------------------------------
' DATE:             05/02/2011
' DESCRIPTION:      Closes The Data File For The Current File Handle.
'---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.FileHandle <> 0 Then
        Close #Me.FileHandle
    End If
    Me.FileHandle = Me.FileHandle - 1
    If Me.FileHandle < 0 Then Me.FileHandle = 0
    Me.FileStatus = FileClosed
End Sub

Public Sub Close_All_File_Handles()
'---------------------------------------------------------------------------------------------------------------------------------------------
' DATE:             05/02/2011
' DESCRIPTION:      Closes The Data File For All File Handles.
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim iFileCounter    As Integer
    
    If Me.FileHandle > 1 Then
        For iFileCounter = Me.FileHandle To 1 Step -1
            Close #iFileCounter
        Next iFileCounter
    ElseIf Me.FileHandle = 1 Then
        Close #Me.FileHandle
    End If
    Me.FileHandle = 0
    Me.FileStatus = FileClosed
End Sub

Public Function Check_Folder(pFolderName As String, Optional ByVal pForceCreate As enmFolderCreation) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
' DATE:             07/03/2011
' DESCRIPTION:      Checks For An Existing Folder.
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sTempFolder As String
    
    sTempFolder = Dir(pFolderName)
    
    If sTempFolder <> "" Then
        Check_Folder = True
    Else
        If IsMissing(pForceCreate) Or pForceCreate = DontCreate Then
            Check_Folder = False
        Else
            Check_Folder = Create_Folder(pFolderName)
         End If
    End If
End Function

Private Function Create_Folder(pFolderName) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
' DATE:             07/03/2011
' DESCRIPTION:      Creates A Folder.
' CALLERS:          CheckFolder
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sTempFolder As String
    
    sTempFolder = pFolderName & IIf(Right$(pFolderName, 1) = "\", "", "\")

    If MakeSureDirectoryPathExists(sTempFolder) <> 0 Then
        Create_Folder = True
    End If
End Function
 
