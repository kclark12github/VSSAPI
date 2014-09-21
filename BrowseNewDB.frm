VERSION 5.00
Begin VB.Form frmBrowseNewDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Database"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   Icon            =   "BrowseNewDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   5
      Top             =   375
      Width           =   2160
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2220
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2445
      TabIndex        =   1
      Top             =   810
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2445
      TabIndex        =   0
      Top             =   345
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Drives:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2415
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folders:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   555
   End
End
Attribute VB_Name = "frmBrowseNewDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewDBPath As String

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    If frmAnalyze.Visible = True Then
        BackupFolder = Dir1.Path
        frmAnalyze.lblCurrentFolder.Caption = "Backup folder: " + BackupFolder
        If Dir(BackupFolder + "\*.*") = "" Then
            frmAnalyze.lblBackupStatus.Caption = "Backup folder is: " + "Empty"
        Else
            frmAnalyze.lblBackupStatus.Caption = "Backup folder is: " + "Full"
        End If
    Else
        frmNewDatabase.txtDBPath.Text = Dir1.Path
    End If
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub Dir1_Change()
    
    ' Enable OK Command button as appropriate
    If Dir1.Path <> "" Then
        NewDBPath = Dir1.Path
        cmdOK.Enabled = True
    End If

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
  
  ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

End Sub
