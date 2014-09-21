VERSION 5.00
Begin VB.Form frmAnalyze 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyze Database"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "Analyze.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmBackup 
      Caption         =   "BackUp folder"
      Height          =   1365
      Left            =   75
      TabIndex        =   16
      Top             =   4035
      Width           =   4395
      Begin VB.CommandButton cmdClear 
         Caption         =   "E&mpty Folder"
         Height          =   330
         Left            =   2943
         TabIndex        =   11
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set &Folder"
         Height          =   330
         Left            =   476
         TabIndex        =   10
         Top             =   915
         Width           =   1185
      End
      Begin VB.Label lblBackupStatus 
         AutoSize        =   -1  'True
         Caption         =   "Backup folder is: "
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label lblCurrentFolder 
         AutoSize        =   -1  'True
         Caption         =   "Backup folder: "
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   315
         Width           =   1080
      End
   End
   Begin VB.Frame frmVerbose 
      Caption         =   "Output Verbosity"
      Height          =   1365
      Left            =   105
      TabIndex        =   15
      Top             =   2595
      Width           =   4395
      Begin VB.OptionButton optV2 
         Caption         =   "S&how only significant errors"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   475
         Width           =   3780
      End
      Begin VB.OptionButton optV3 
         Caption         =   "Show all errors and &inconsistencies"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   725
         Width           =   3780
      End
      Begin VB.OptionButton optV4 
         Caption         =   "Show e&rrors, inconsistencies, and informational notes"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   975
         Width           =   4155
      End
      Begin VB.OptionButton optV1 
         Caption         =   "Show only cri&tical errors"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   3780
      End
   End
   Begin VB.Frame frmAnalyzeOptions 
      Caption         =   "Analyze options"
      Height          =   2400
      Left            =   105
      TabIndex        =   14
      Top             =   60
      Width           =   4395
      Begin VB.CheckBox chkExit 
         Caption         =   "&When the analysis is complete the program exits"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1860
         Width           =   3750
      End
      Begin VB.CheckBox chkFix 
         Caption         =   "&Automatically fix files with corruptions"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1590
         Width           =   3750
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "&Delete unused files"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3750
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "&Show Analyze help"
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4035
      End
      Begin VB.CheckBox chkLock 
         Caption         =   "Do &not lock the database when analyzing"
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   495
         Width           =   3540
      End
      Begin VB.CheckBox chkCompress 
         Caption         =   "&Compress unused space"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1050
         Width           =   3750
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Canc&el"
      Height          =   345
      Left            =   2467
      TabIndex        =   13
      Top             =   5535
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   982
      TabIndex        =   12
      Top             =   5535
      Width           =   1155
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Verbosity As String
Dim VSSPath As String

Private Sub chkHelp_Click()

    ' Disable other options as appropriate
    If chkHelp.Value = 1 Then
        chkCompress.Enabled = False
        chkLock.Enabled = False
        chkDelete.Enabled = False
        chkFix.Enabled = False
        chkExit.Enabled = False
    Else
        chkLock.Enabled = True
        chkExit.Enabled = True
        If Not LoggedOn Then
            If chkLock.Value = 1 Then
                chkCompress.Enabled = False
                chkDelete.Enabled = False
                chkFix.Enabled = False
            Else
                chkCompress.Enabled = True
                chkDelete.Enabled = True
                chkFix.Enabled = True
            End If
        End If
    End If

End Sub

Private Sub chkLock_Click()
    
    ' Disable other options as appropriate
    If chkLock.Value = 1 Then
        chkCompress.Enabled = False
        chkDelete.Enabled = False
        chkFix.Enabled = False
        chkHelp.Enabled = False
        cmdSet.Enabled = False
    Else
        If Not LoggedOn Then
            chkCompress.Enabled = True
            chkDelete.Enabled = True
            chkFix.Enabled = True
            cmdSet.Enabled = True
        End If
        chkHelp.Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdClear_Click()

    Dim Response As Long
    Dim RetVal As Long

    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    RetVal = MsgBox("This will delete all files in the folder '" + BackupFolder + "'. Are you sure?", vbYesNo + vbQuestion, AppTitle)
    If RetVal = vbYes Then
    
        ' Delete files
        Kill (BackupFolder + "\*.*")
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
    
        If Err <> 53 Then
            Response = MsgBox("Unable to empty backup folder." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        ElseIf Err = 53 Then
            Response = MsgBox("Backup folder is already empty.", vbExclamation, AppTitle)

        End If
        Err.Clear
    Else
        lblBackupStatus.Caption = "Backup folder is: " + "Empty"
    End If
    
End Sub

Private Sub cmdOK_Click()

    Dim Response As Long
    Dim HelpFlag As String
    Dim LockFlag As String
    Dim CompressFlag As String
    Dim DeleteFlag As String
    Dim FixFlag As String
    Dim ExitFlag As String
    Dim BackUpFlag As String
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Set Analyze switches
    If chkHelp.Enabled = True Then
        If chkHelp.Value = 1 Then
            HelpFlag = " -h "
            RetVal = Shell(VSSPath + "WIN32\ANALYZE.EXE" + HelpFlag, vbNormalFocus)
            Exit Sub
        Else
            HelpFlag = ""
        End If
    End If
    If chkLock.Value = 1 Then
        LockFlag = " -x "
    Else
        LockFlag = ""
        
        ' Warn user that we will shut down during analysis
        Response = MsgBox("This application will close while Analyze is being run. Continue?", vbYesNo + vbQuestion, AppTitle)
        If Response = vbYes Then
        
            ' Disconnect from the VSS Database Object
            Set objVSSDatabase = Nothing
            Set objVSSProject = Nothing
            Set objVSSVersion = Nothing
            Set objVSSCheckOut = Nothing
            Set objVSSObject = Nothing
        Else
            Exit Sub
        End If
    End If
    If chkExit.Value = 1 Then
        ExitFlag = " -i- "
    Else
        ExitFlag = ""
    End If
    If chkDelete.Enabled = True Then
        If chkDelete.Value = 1 Then
            DeleteFlag = " -d "
        Else
            DeleteFlag = ""
        End If
    Else
        DeleteFlag = ""
    End If
    If chkFix.Enabled = True Then
        If chkFix.Value = 1 Then
            FixFlag = " -f "
        Else
            FixFlag = ""
        End If
    Else
        FixFlag = ""
    End If
    If chkCompress.Enabled = True Then
        If chkCompress.Value = 1 Then
            CompressFlag = " -c "
        Else
            CompressFlag = ""
        End If
    Else
        CompressFlag = ""
    End If
    BackUpFlag = "-b" + BackupFolder + " "


    ' Launch Analyze
    RetVal = Shell(VSSPath + "WIN32\ANALYZE.EXE " + VSSPath + "Data" + FixFlag + HelpFlag + LockFlag + DeleteFlag + CompressFlag + ExitFlag + Verbosity + BackUpFlag, vbNormalFocus)
    
    ' Close application as needed
    If LockFlag = "" And Response = vbYes Then
        
        End
    
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
        If Err = 53 Then
            Response = MsgBox(" This command may only be run from a database that has a WIN32 directory containing" + vbCrLf + "ANALYZE.EXE. Please select an appropriate database and try again.", vbCritical, AppTitle)
            Unload Me
        Else
            Response = MsgBox("Cannot Analyze database." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If

    End If
    
End Sub

Private Sub cmdSet_Click()

    frmBrowseNewDB.Caption = "Set Backup Folder"
    frmBrowseNewDB.Show 1
    
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Get the VSS path
    VSSPath = GetVSSPath(objVSSDatabase.SrcSafeIni)
    If Right(VSSPath, 1) <> "\" Then VSSPath = VSSPath + "\"
    
    ' Set default verbosity
    Verbosity = " -V1 "
    
    ' Set default backupfolder
    BackupFolder = VSSPath + "Data\Backup"
    lblCurrentFolder.Caption = lblCurrentFolder.Caption + BackupFolder
    If Dir(BackupFolder, vbDirectory) = "" Then
        MkDir BackupFolder
        lblBackupStatus.Caption = lblBackupStatus.Caption + "Empty"
    Else
        If Dir(BackupFolder + "\*.*") = "" Then
            lblBackupStatus.Caption = lblBackupStatus.Caption + "Empty"
        Else
            lblBackupStatus.Caption = lblBackupStatus.Caption + "Full"
        End If
    End If
    
End Sub

Private Sub optV1_Click()
    
    Verbosity = " -v1"
    
End Sub

Private Sub optV2_Click()
    
    Verbosity = " -v2"
    
End Sub

Private Sub optV3_Click()

    Verbosity = " -v3"
    
End Sub

Private Sub optV4_Click()

    Verbosity = " -v4"
    
End Sub
