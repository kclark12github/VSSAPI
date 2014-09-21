VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5400
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   9525
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Options.frx":014A
      Tab(0).ControlCount=   5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCurrentProject"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSrcSafeini"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUser"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDatabaseName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      TabCaption(1)   =   "Diff/View"
      TabPicture(1)   =   "Options.frx":0166
      Tab(1).ControlCount=   4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkShowLineNumbers"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "chkUseLineMarker"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      TabCaption(2)   =   "Warnings"
      TabPicture(2)   =   "Options.frx":0182
      Tab(2).ControlCount=   8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkWarnExit"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "chkWarnDestroy"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "chkWarnDelete"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "chkWarnPurge"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "chkWarnCheckedout"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "chkWarnUndoCheckout"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "lblWarnings"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblInfo"
      Tab(2).Control(7).Enabled=   0   'False
      TabCaption(3)   =   "Admin Options"
      TabPicture(3)   =   "Options.frx":019E
      Tab(3).ControlCount=   5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "ListView1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblUsers"
      Tab(3).Control(4).Enabled=   0   'False
      Begin VB.CheckBox chkShowLineNumbers 
         Caption         =   "&Show line numbers (OCX control only)"
         Height          =   330
         Left            =   -74640
         TabIndex        =   53
         Top             =   465
         Width           =   3870
      End
      Begin VB.CheckBox chkUseLineMarker 
         Caption         =   "&Use line marker (OCX control only)"
         Height          =   330
         Left            =   -74640
         TabIndex        =   52
         Top             =   720
         Width           =   3870
      End
      Begin VB.Frame Frame6 
         Caption         =   "When Showing File Difference"
         Height          =   2955
         Left            =   -74820
         TabIndex        =   0
         Top             =   2250
         Width           =   4980
         Begin VB.TextBox txtContext 
            Height          =   285
            Left            =   3195
            TabIndex        =   17
            Top             =   2085
            Width           =   1230
         End
         Begin VB.CheckBox chkShowContext 
            Caption         =   "Sho&w context"
            Height          =   225
            Left            =   2550
            TabIndex        =   16
            Top             =   1785
            Width           =   1785
         End
         Begin VB.CheckBox chkIgnoreCase 
            Caption         =   "Ig&nore case (OCX control only)"
            Height          =   330
            Left            =   180
            TabIndex        =   10
            Top             =   810
            Width           =   3870
         End
         Begin VB.CheckBox chkIgnoreOS 
            Caption         =   "Ignore OS dif&ferences (OCX control only)"
            Height          =   330
            Left            =   180
            TabIndex        =   11
            Top             =   1095
            Width           =   3870
         End
         Begin VB.OptionButton optSS 
            Caption         =   "So&urceSafe"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   2250
            Width           =   1350
         End
         Begin VB.OptionButton optUnix 
            Caption         =   "&Unix"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   2490
            Width           =   1020
         End
         Begin VB.OptionButton optVisual 
            Caption         =   "&Visual"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   2025
            Width           =   1020
         End
         Begin VB.CheckBox chkModal 
            Caption         =   "&Diff window is modal (OCX control only)"
            Height          =   330
            Left            =   180
            TabIndex        =   8
            Top             =   240
            Width           =   3870
         End
         Begin VB.CheckBox chkIgnoreWhite 
            Caption         =   "&Ignore white space (OCX control only)"
            Height          =   330
            Left            =   180
            TabIndex        =   9
            Top             =   525
            Width           =   3870
         End
         Begin VB.CheckBox chkDiffMethod 
            Caption         =   "Use O&CX Control"
            Height          =   315
            Left            =   180
            TabIndex        =   12
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label lblLines 
            AutoSize        =   -1  'True
            Caption         =   "Lines:"
            Height          =   195
            Left            =   2550
            TabIndex        =   55
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label lblFormat 
            AutoSize        =   -1  'True
            Caption         =   "Format (OCX control only):"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   1785
            Width           =   1845
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "When Viewing\Editing Files"
         Height          =   1020
         Left            =   -74820
         TabIndex        =   50
         Top             =   1170
         Width           =   4980
         Begin VB.CheckBox chkViewMethod 
            Caption         =   "Use &OCX Control"
            Height          =   315
            Left            =   180
            TabIndex        =   6
            Top             =   270
            Width           =   1785
         End
         Begin VB.TextBox txtTabStop 
            Height          =   285
            Left            =   2700
            TabIndex        =   7
            Top             =   630
            Width           =   570
         End
         Begin VB.Label lblTabStop 
            AutoSize        =   -1  'True
            Caption         =   "Tab stop width (OCX control only)"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   660
            Width           =   2370
         End
      End
      Begin VB.CheckBox chkWarnExit 
         Caption         =   "&Exit OLE Automation Sample"
         Height          =   345
         Left            =   -74685
         TabIndex        =   47
         Top             =   3015
         Width           =   3090
      End
      Begin VB.CheckBox chkWarnDestroy 
         Caption         =   "&Destroy a file or project"
         Height          =   345
         Left            =   -74685
         TabIndex        =   46
         Top             =   1860
         Width           =   3090
      End
      Begin VB.CheckBox chkWarnDelete 
         Caption         =   "De&lete a file or project"
         Height          =   345
         Left            =   -74685
         TabIndex        =   45
         Top             =   1485
         Width           =   3090
      End
      Begin VB.CheckBox chkWarnPurge 
         Caption         =   "&Purge a file or project"
         Height          =   345
         Left            =   -74685
         TabIndex        =   44
         Top             =   2250
         Width           =   3090
      End
      Begin VB.CheckBox chkWarnCheckedout 
         Caption         =   "&Check out an already checked out file"
         Height          =   345
         Left            =   -74685
         TabIndex        =   43
         Top             =   2625
         Width           =   3090
      End
      Begin VB.CheckBox chkWarnUndoCheckout 
         Caption         =   "&Undo checkout of a modified file"
         Height          =   345
         Left            =   -74685
         TabIndex        =   42
         Top             =   1095
         Width           =   3090
      End
      Begin VB.Frame Frame3 
         Caption         =   "User Options"
         Height          =   1770
         Left            =   -72885
         TabIndex        =   36
         Top             =   555
         Width           =   1365
         Begin VB.CommandButton cmdAddUser 
            Caption         =   "&Add User..."
            Height          =   360
            Left            =   90
            TabIndex        =   39
            Top             =   330
            Width           =   1185
         End
         Begin VB.CommandButton cmdDeleteUser 
            Caption         =   "&Delete User"
            Height          =   360
            Left            =   90
            TabIndex        =   38
            Top             =   780
            Width           =   1185
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit User..."
            Height          =   360
            Left            =   90
            TabIndex        =   37
            Top             =   1230
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Security"
         Height          =   1770
         Left            =   -74865
         TabIndex        =   32
         Top             =   555
         Width           =   1950
         Begin VB.CommandButton cmdRights 
            Caption         =   "&User Rights..."
            Height          =   360
            Left            =   165
            TabIndex        =   35
            Top             =   1170
            Width           =   1395
         End
         Begin VB.CommandButton cmdDefaultRights 
            Caption         =   "Default &Rights..."
            Height          =   360
            Left            =   165
            TabIndex        =   34
            Top             =   690
            Width           =   1395
         End
         Begin VB.CheckBox chkProjectRights 
            Caption         =   "&Enable Project Rights"
            Height          =   300
            Left            =   60
            TabIndex        =   33
            Top             =   270
            Width           =   1830
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Maintenance"
         Height          =   1770
         Left            =   -71445
         TabIndex        =   28
         Top             =   555
         Width           =   1680
         Begin VB.CommandButton cmdNewDatabase 
            Caption         =   "&Create Database..."
            Height          =   360
            Left            =   75
            TabIndex        =   31
            Top             =   330
            Width           =   1530
         End
         Begin VB.CommandButton cmdFormat 
            Caption         =   "Change &Format..."
            Height          =   360
            Left            =   75
            TabIndex        =   30
            Top             =   780
            Width           =   1530
         End
         Begin VB.CommandButton cmdAnalyze 
            Caption         =   "Analy&ze..."
            Height          =   360
            Left            =   75
            TabIndex        =   29
            Top             =   1230
            Width           =   1530
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "User Options"
         Height          =   2385
         Left            =   323
         TabIndex        =   24
         Top             =   2040
         Width           =   4695
         Begin VB.ComboBox cboMerge 
            Height          =   315
            Left            =   2190
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   1845
            Width           =   2265
         End
         Begin VB.ComboBox cboDoubleClick 
            Height          =   315
            Left            =   2565
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   1455
            Width           =   1890
         End
         Begin VB.ComboBox cboCheckIn 
            Height          =   315
            Left            =   2565
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   1065
            Width           =   1890
         End
         Begin VB.ComboBox cboEOL 
            Height          =   315
            Left            =   2565
            TabIndex        =   2
            Text            =   "Combo1"
            Top             =   675
            Width           =   1890
         End
         Begin VB.CheckBox chkBuildTree 
            Caption         =   "&Build Tree (override working folders)"
            Height          =   345
            Left            =   165
            TabIndex        =   1
            Top             =   240
            Width           =   3090
         End
         Begin VB.Label lblMerge 
            AutoSize        =   -1  'True
            Caption         =   "Use Visual Merge:"
            Height          =   195
            Left            =   165
            TabIndex        =   56
            Top             =   1905
            Width           =   1290
         End
         Begin VB.Label lblDoubleClick 
            AutoSize        =   -1  'True
            Caption         =   "Double Click on a file:"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   1515
            Width           =   1545
         End
         Begin VB.Label lblCheckInFiles 
            AutoSize        =   -1  'True
            Caption         =   "Check In unchanged files:"
            Height          =   195
            Left            =   165
            TabIndex        =   26
            Top             =   1125
            Width           =   1860
         End
         Begin VB.Label lblEOL 
            AutoSize        =   -1  'True
            Caption         =   "End-of-line charater for text files:"
            Height          =   195
            Left            =   165
            TabIndex        =   25
            Top             =   735
            Width           =   2265
         End
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   2610
         Left            =   -74895
         TabIndex        =   40
         Top             =   2670
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   4604
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Name"
            Object.Tag             =   ""
            Text            =   "User Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   "Rights"
            Object.Tag             =   ""
            Text            =   "Rights"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   "LoggedIn"
            Object.Tag             =   ""
            Text            =   "Logged In"
            Object.Width           =   2205
         EndProperty
      End
      Begin VB.Label lblWarnings 
         Caption         =   "Display warning for these comands:"
         Height          =   285
         Left            =   -74790
         TabIndex        =   49
         Top             =   690
         Width           =   2700
      End
      Begin VB.Label lblInfo 
         Caption         =   $"Options.frx":01BA
         Height          =   420
         Left            =   -74685
         TabIndex        =   48
         Top             =   3705
         Width           =   4800
      End
      Begin VB.Label lblUsers 
         AutoSize        =   -1  'True
         Caption         =   "Users:"
         Height          =   195
         Left            =   -74895
         TabIndex        =   41
         Top             =   2460
         Width           =   450
      End
      Begin VB.Label lblDatabaseName 
         AutoSize        =   -1  'True
         Caption         =   "Database Name:"
         Height          =   195
         Left            =   323
         TabIndex        =   23
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "Logged on as:"
         Height          =   195
         Left            =   323
         TabIndex        =   22
         Top             =   705
         Width           =   1020
      End
      Begin VB.Label lblSrcSafeini 
         AutoSize        =   -1  'True
         Caption         =   "SourceSafe.ini location:"
         Height          =   195
         Left            =   323
         TabIndex        =   21
         Top             =   1275
         Width           =   1680
      End
      Begin VB.Label lblCurrentProject 
         AutoSize        =   -1  'True
         Caption         =   "Current Project:"
         Height          =   195
         Left            =   323
         TabIndex        =   20
         Top             =   990
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   4185
      TabIndex        =   18
      Top             =   5460
      Width           =   1140
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDoubleClick_Click()
    
    DoubleClickFile = cboDoubleClick.Text
    
End Sub

Private Sub chkModal_Click()
    
    ' Set Diff Method
    If chkModal.Value = 1 Then
        frmMain.Diff1.VisualDiffModal = True
    Else
        frmMain.Diff1.VisualDiffModal = False
    End If
    
End Sub

Private Sub chkBuildTree_Click()
    
    If chkBuildTree.Value = 0 Then
        ForceDirFlag = VSSFLAG_FORCEDIRYES
    Else
        ForceDirFlag = VSSFLAG_FORCEDIRNO
    End If

End Sub

Private Sub chkIncludeFiles_Click()

    If chkIncludeFiles.Value = 1 Then
        IgnoreFiles = 0
    Else
        IgnoreFiles = VSSFLAG_HISTIGNOREFILES
    End If
    
End Sub

Private Sub chkCloseWarning_Click()

    If chkCloseWarning.Value = 0 Then
        ExitWarning = False
    Else
        ExitWarning = True
    End If
End Sub

Private Sub chkDiffMethod_Click()

    ' Set Diff Method
    If chkDiffMethod.Value = 1 Then
        DiffMethod = True
        chkModal.Enabled = True
        chkIgnoreWhite.Enabled = True
        chkIgnoreCase.Enabled = True
        chkIgnoreOS.Enabled = True
        optVisual.Enabled = True
        optSS.Enabled = True
        optUnix.Enabled = True
        chkShowContext.Enabled = True
        If ShowContext Then
            lblLines.Enabled = True
            txtContext.Enabled = True
        End If
    Else
        DiffMethod = False
        chkModal.Enabled = False
        chkIgnoreWhite.Enabled = False
        chkIgnoreCase.Enabled = False
        chkIgnoreOS.Enabled = False
        optVisual.Enabled = False
        optSS.Enabled = False
        optUnix.Enabled = False
        chkShowContext.Enabled = False
        lblLines.Enabled = False
        txtContext.Enabled = False
    End If

End Sub

Private Sub chkProjectRights_Click()

    If chkProjectRights.Value = 1 Then
        objVSSDatabase.ProjectRightsEnabled = True
        cmdRights.Enabled = True
        cmdDefaultRights.Enabled = True
    Else
        objVSSDatabase.ProjectRightsEnabled = False
        cmdRights.Enabled = False
        cmdDefaultRights.Enabled = False
    End If
    
End Sub

Private Sub chkShowContext_Click()
    
    If chkShowContext.Value = 1 Then
        ShowContext = True
        txtContext.Enabled = True
        lblLines.Enabled = True
    Else
        ShowContext = False
        txtContext.Enabled = False
        lblLines.Enabled = False
    End If
    
End Sub

Private Sub chkShowLineNumbers_Click()

    If chkShowLineNumbers.Value = 1 Then
        frmMain.Viewer1.ShowLineNumbers = True
    Else
        frmMain.Viewer1.ShowLineNumbers = False
    End If

End Sub

Private Sub chkUseLineMarker_Click()

    If chkUseLineMarker.Value = 1 Then
        frmMain.Viewer1.UseLineMarker = True
    Else
        frmMain.Viewer1.UseLineMarker = False
    End If
    
End Sub

Private Sub chkViewMethod_Click()

    ' Set View Method
    If chkViewMethod.Value = 1 Then
        ViewMethod = True
        chkShowLineNumbers.Enabled = True
        chkUseLineMarker.Enabled = True
        txtTabStop.Enabled = True
        lblTabStop.Enabled = True
    Else
        ViewMethod = False
        chkShowLineNumbers.Enabled = False
        chkUseLineMarker.Enabled = False
        txtTabStop.Enabled = False
        lblTabStop.Enabled = False
    End If

End Sub

Private Sub chkWarnCheckedout_Click()
    
    If chkWarnCheckedout.Value = 0 Then
        WarnCheckOut = False
    Else
        WarnCheckOut = True
    End If
    
End Sub

Private Sub chkWarnDelete_Click()

    If chkWarnDelete.Value = 0 Then
        WarnDestroy = False
    Else
        WarnDelete = True
    End If
    
End Sub

Private Sub chkWarnDestroy_Click()

    If chkWarnDestroy.Value = 0 Then
        WarnDestroy = False
    Else
        WarnDestroy = True
    End If
    
End Sub

Private Sub chkWarnExit_Click()

    If chkWarnExit.Value = 0 Then
        WarnExit = False
    Else
        WarnExit = True
    End If

End Sub

Private Sub chkWarnPurge_Click()

    If chkWarnPurge.Value = 0 Then
        WarnPurge = False
    Else
        WarnPurge = True
    End If
    
End Sub

Private Sub chkWarnUndoCheckout_Click()
    
    If chkWarnUndoCheckout.Value = 0 Then
        WarnUndoCheckOut = False
    Else
        WarnUndoCheckOut = True
    End If
    
End Sub

Private Sub cmdAddUser_Click()

    frmAddUser.Show 1
    
End Sub

Private Sub cmdAnalyze_Click()
    
    Dim Response As Long
    
    ' Check for other users logged on
    If LoggedOn Then
        Response = MsgBox("There are other users logged on to the database." + vbCrLf + "Certain Analyze options will be disabled. Continue?", vbYesNo + vbQuestion, AppTitle)
        If Response = vbYes Then
            frmAnalyze.chkCompress.Enabled = False
            frmAnalyze.chkDelete.Enabled = False
            frmAnalyze.chkFix.Enabled = False
            frmAnalyze.chkLock.Value = 1
            frmAnalyze.Show 1
        End If
    Else
        frmAnalyze.Show 1
    End If
    
End Sub

Private Sub cmdDefaultRights_Click()

    frmDefaultRights.Show 1

End Sub

Private Sub cmdDeleteUser_Click()

    Dim objUserToDelete As VSSUser
    Dim Response As Long
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Instantiate the user object
    Set objUserToDelete = objVSSDatabase.User(ListView1.SelectedItem)
    
    ' Delete user
    objUserToDelete.Delete
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
    
        Response = MsgBox("Unable to delete user." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    Else
        ListView1.ListItems.Remove (objUserToDelete.Name)
    End If

End Sub

Private Sub cmdEdit_Click()

    Dim Response As Long
    
    ' Initialize variables
    Set CurrentUser = objVSSDatabase.User(ListView1.SelectedItem.Text)
    
    frmEditUser.Caption = "Edit User " + CurrentUser.Name
    If UCase(CurrentUser.Name) = "ADMIN" Then
        Response = MsgBox("Sorry, you cannot modify the rights for user " + CurrentUser.Name + "." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    Else
        If CurrentUser.ReadOnly = True Then frmEditUser.chkReadOnly.Value = 1
        frmEditUser.txtName.Text = CurrentUser.Name
        frmEditUser.Show 1
    End If

End Sub

Private Sub cmdFormat_Click()
    
    frmFormat.Show 1

End Sub

Private Sub cmdNewDatabase_Click()

    frmNewDatabase.Show 1

End Sub

Private Sub cmdOK_Click()

    Dim Response As Long
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Set Context
    If ShowContext Then
        frmMain.Diff1.DiffContext = Val(txtContext.Text)
    Else
        frmMain.Diff1.DiffContext = -1
    End If
    
    ' Set Viewer TabStop
    frmMain.Viewer1.ViewerTab = Val(txtTabStop.Text)
    
    ' Set Visual Merge
    Select Case cboMerge
        
        Case "Yes"
            
            frmMain.VisMerge1.VisualMerge = "Yes"
        
        Case "No"
        
            frmMain.VisMerge1.VisualMerge = "No"
        
        Case "Only if there are conflicts"
        
            frmMain.VisMerge1.VisualMerge = "Conflicts"
        
    End Select
    
    ' Set EOL Variable
    Select Case cboEOL.Text
    
        Case "CRLF"
        
            EOLFlag = VSSFLAG_EOLCRLF
            
        Case "CR"
            
            EOLFlag = VSSFLAG_EOLCR
        
        Case "LF"
        
            EOLFlag = VSSFLAG_EOLLF
           
    End Select
    
    ' Set Check In Unchanged Files Flag
    If cboCheckIn.Text = "Check In" Then
        CheckInUnchangedFlag = VSSFLAG_UPDUPDATE
    ElseIf cboCheckIn.Text = "Undo Check Out" Then
        CheckInUnchangedFlag = VSSFLAG_UPDUNCH
    End If
    
    ' Set Diff Ignore Switch
    If chkIgnoreWhite.Value = 0 And chkIgnoreCase.Value = 0 And chkIgnoreOS.Value = 1 Then
        frmMain.Diff1.DiffIgnore = "w-c-e"
    ElseIf chkIgnoreWhite.Value = 0 And chkIgnoreCase.Value = 1 And chkIgnoreOS.Value = 1 Then
        frmMain.Diff1.DiffIgnore = "w-ce"
    ElseIf chkIgnoreWhite.Value = 1 And chkIgnoreCase.Value = 1 And chkIgnoreOS.Value = 1 Then
        frmMain.Diff1.DiffIgnore = "wce"
    ElseIf chkIgnoreWhite.Value = 1 And chkIgnoreCase.Value = 0 And chkIgnoreOS.Value = 1 Then
        frmMain.Diff1.DiffIgnore = "wc-e"
    ElseIf chkIgnoreWhite.Value = 0 And chkIgnoreCase.Value = 1 And chkIgnoreOS.Value = 0 Then
        frmMain.Diff1.DiffIgnore = "w-ce-"
    ElseIf chkIgnoreWhite.Value = 1 And chkIgnoreCase.Value = 1 And chkIgnoreOS.Value = 0 Then
        frmMain.Diff1.DiffIgnore = "wce-"
    ElseIf chkIgnoreWhite.Value = 1 And chkIgnoreCase.Value = 0 And chkIgnoreOS.Value = 0 Then
        frmMain.Diff1.DiffIgnore = "wc-e-"
    ElseIf chkIgnoreWhite.Value = 0 And chkIgnoreCase.Value = 0 And chkIgnoreOS.Value = 0 Then
        frmMain.Diff1.DiffIgnore = "w-c-e-"
    End If
    
    VSSUser = ""

    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:

        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
    
    Else
    
        ' Close Form
        Unload Me
        
    End If
    
End Sub

Private Sub cmdRights_Click()

    Dim Response As Long
    
    ' Initialize variables
    Set CurrentUser = objVSSDatabase.User(ListView1.SelectedItem.Text)
    
    frmRights.Caption = "Assign Project Rights for user " + CurrentUser.Name
    If UCase(CurrentUser.Name) = "ADMIN" Then
        Response = MsgBox("Sorry, you cannot modify the rights for user " + CurrentUser.Name + "." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    Else
        frmRights.Show 1
    End If

End Sub

Private Sub Form_Load()

    Dim objUserListItem As ListItem
    Dim UserRights As String
    Dim DatabaseID As String
    Dim LoggedPath As String
    Dim FileAttr As Integer
    Dim Response As Long
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    LoggedOn = False
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Set user personal settings
    If WarnExit Then chkWarnExit.Value = 1
    If WarnDelete Then chkWarnDelete.Value = 1
    If WarnDestroy Then chkWarnDestroy.Value = 1
    If WarnPurge Then chkWarnPurge.Value = 1
    If WarnCheckOut Then chkWarnCheckedout.Value = 1
    If WarnUndoCheckOut Then chkWarnUndoCheckout.Value = 1
    
    ' Set Diff\View tab settings
    txtTabStop.Text = Str(frmMain.Viewer1.ViewerTab)
    If frmMain.Viewer1.ShowLineNumbers = True Then
        chkShowLineNumbers.Value = 1
    Else
        chkShowLineNumbers.Value = 0
    End If
    If frmMain.Viewer1.UseLineMarker = True Then
        chkUseLineMarker.Value = 1
    Else
        chkUseLineMarker.Value = 0
    End If
    If ViewMethod Then
        chkViewMethod.Value = 1
    Else
        chkShowLineNumbers.Enabled = False
        chkUseLineMarker.Enabled = False
        txtTabStop.Enabled = False
        lblTabStop.Enabled = False
    End If
    
    ' Set Context
    If ShowContext Then
        chkShowContext.Value = 1
        txtContext.Text = frmMain.Diff1.DiffContext
    Else
        chkShowContext.Value = 0
        txtContext.Text = 3
        txtContext.Enabled = False
        lblLines.Enabled = False
    End If
    
    ' Set Diff Method properties
    If DiffMethod Then
        chkDiffMethod.Value = 1
    Else
        chkModal.Enabled = False
        chkIgnoreWhite.Enabled = False
        chkIgnoreCase.Enabled = False
        chkIgnoreOS.Enabled = False
        optVisual.Enabled = False
        optSS.Enabled = False
        optUnix.Enabled = False
        chkShowContext.Enabled = False
        lblLines.Enabled = False
        txtContext.Enabled = False
    End If
    If frmMain.Diff1.VisualDiffModal Then chkModal.Value = 1
    Select Case UCase(frmMain.Diff1.DiffIgnore)
    
        Case "W-C-E", "WE"
            chkIgnoreWhite.Value = 0
            chkIgnoreCase.Value = 0
            chkIgnoreOS.Value = 1
        Case "W-CE"
            chkIgnoreWhite.Value = 0
            chkIgnoreCase.Value = 1
            chkIgnoreOS.Value = 1
        Case "WCE"
            chkIgnoreWhite.Value = 1
            chkIgnoreCase.Value = 1
            chkIgnoreOS.Value = 1
        Case "WC-E"
            chkIgnoreWhite.Value = 1
            chkIgnoreCase.Value = 0
            chkIgnoreOS.Value = 1
        Case "W-CE-"
            chkIgnoreWhite.Value = 0
            chkIgnoreCase.Value = 1
            chkIgnoreOS.Value = 0
        Case "WCE-"
            chkIgnoreWhite.Value = 1
            chkIgnoreCase.Value = 1
            chkIgnoreOS.Value = 0
        Case "WC-E-"
            chkIgnoreWhite.Value = 1
            chkIgnoreCase.Value = 0
            chkIgnoreOS.Value = 0
        Case "W-C-E-"
            chkIgnoreWhite.Value = 0
            chkIgnoreCase.Value = 0
            chkIgnoreOS.Value = 0
    End Select
    Select Case UCase(frmMain.Diff1.DiffFormat)
        
        Case "VISUAL"
            
            optVisual.Value = True
            
        Case "SS"
            
            optSS.Value = True
            
        Case "UNIX"
            
            optUnix.Value = True
    
    End Select

    ' Get Loggedin path
    LoggedPath = GetDirPath(objVSSDatabase.SrcSafeIni) + "\Data\Loggedin\"
    
    ' Initialize dialog
    SSTab1.Tab = 0
    If UserName = "Admin" Then
    
        ' Set Enable Rights CheckBox
        If objVSSDatabase.ProjectRightsEnabled = True Then
            chkProjectRights.Value = 1
        Else
            chkProjectRights.Value = 0
            cmdRights.Enabled = False
            cmdDefaultRights.Enabled = False
        End If
        
        ' Load user list
        ListView1.ColumnHeaders(1).Text = "User Name:"
        ListView1.ColumnHeaders(2).Text = "Rights:"
        ListView1.ColumnHeaders(3).Text = "Logged In:"
        For Each User In objVSSDatabase.Users
            
            ' Retrieve and display rights
            Set objUserListItem = ListView1.ListItems.Add(, User.Name, User.Name)
            If User.ReadOnly = True Then
                UserRights = "Read-Only"
            Else
                UserRights = "Read-Write"
            End If
            objUserListItem.SubItems(1) = UserRights
            
            ' Retrieve and display Logged In status
            If Dir(LoggedPath + User.Name + ".log") <> "" Then
                objUserListItem.SubItems(2) = "Yes"
                If UCase(User.Name) <> "ADMIN" Then
                    LoggedOn = True
                End If
            End If
        Next
        Set ListView1.SelectedItem = ListView1.ListItems(1)
    Else
        SSTab1.TabEnabled(3) = False
    End If
    
    ' Populate User information
    lblCurrentProject.Caption = "Current project: " + objVSSDatabase.CurrentProject
    lblSrcSafeini.Caption = "Srcsafe.ini location: " + objVSSDatabase.SrcSafeIni
    lblUser.Caption = "Logged in as: " + objVSSDatabase.UserName
    If objVSSDatabase.DatabaseName = "" Then
        DatabaseID = "None"
    Else
        DatabaseID = objVSSDatabase.DatabaseName
    End If
    lblDatabaseName.Caption = "Database name: " + DatabaseID
    If ForceDirFlag = VSSFLAG_FORCEDIRNO Then
        chkBuildTree.Value = 1
    ElseIf ForceDirFlag = VSSFLAG_FORCEDIRYES Then
        chkBuildTree.Value = 0
    End If
    If IgnoreFiles = 0 Then chkIncludeFiles = 1
    cboEOL.AddItem ("CRLF")
    cboEOL.AddItem ("CR")
    cboEOL.AddItem ("LF")
    cboEOL.ListIndex = 0
    cboCheckIn.AddItem ("Check In")
    cboCheckIn.AddItem ("Undo Check Out")
    If CheckInUnchangedFlag = VSSFLAG_UPDUPDATE Then
        cboCheckIn.ListIndex = 0
    ElseIf CheckInUnchangedFlag = VSSFLAG_UPDUNCH Then
        cboCheckIn.ListIndex = 1
    End If
    cboDoubleClick.AddItem ("Ask")
    cboDoubleClick.AddItem ("View File")
    cboDoubleClick.AddItem ("Edit File")
    cboDoubleClick.Text = DoubleClickFile
    
    ' Populate Visual Merge Combo
    cboMerge.AddItem ("No")
    cboMerge.AddItem ("Only if there are conflicts")
    cboMerge.AddItem ("Yes")
    Select Case frmMain.VisMerge1.VisualMerge
    
        Case "Conflicts"
            
            cboMerge.ListIndex = 1
        
        Case "Yes"
            
            cboMerge.ListIndex = 2
        
        Case "NO"
            
            cboMerge.ListIndex = 0
            
    End Select
    
ErrHandler:
    If Err <> 0 Then
        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
    
End Sub


Private Sub optSS_Click()

    ' Set Format
    If optSS.Value = True Then frmMain.Diff1.DiffFormat = "ss"
    
End Sub

Private Sub optVisual_Click()

    ' Set Format
    If optVisual.Value = True Then frmMain.Diff1.DiffFormat = "visual"

End Sub

Private Sub optUnix_Click()
    
    ' Set Format
    If optUnix.Value = True Then frmMain.Diff1.DiffFormat = "unix"
    
End Sub


