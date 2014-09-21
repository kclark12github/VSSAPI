VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for SRCSAFE.INI"
   ClientHeight    =   3690
   ClientLeft      =   2235
   ClientTop       =   3480
   ClientWidth     =   4740
   Icon            =   "browse.frx":0000
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3690
   ScaleWidth      =   4740
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Text            =   "srcsafe.ini"
      Top             =   3330
      Width           =   4665
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   3600
      TabIndex        =   4
      Top             =   30
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3600
      TabIndex        =   5
      Top             =   435
      Width           =   1020
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   45
      Pattern         =   "srcsafe.ini"
      TabIndex        =   0
      Top             =   810
      Width           =   2235
   End
   Begin VB.DirListBox Dir1 
      Height          =   1380
      Left            =   2475
      TabIndex        =   1
      Top             =   810
      Width           =   2160
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2475
      TabIndex        =   2
      Top             =   2700
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Files:"
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   450
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folders:"
      Height          =   195
      Left            =   2445
      TabIndex        =   8
      Top             =   450
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Drives:"
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   2355
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "List files of type:"
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   3090
      Width           =   1125
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()

    ' Populate the Login (frmLogon) SRCSAFE.INI text box
    ' with the selected file and then close this form
    Call PopulateSrcSafeini
    
End Sub

Private Sub cmdCancel_Click()

    ' User canceled this form
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    cmdOpen.Enabled = False
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    cmdOpen.Enabled = False
End Sub

Private Sub File1_Click()

    ' If a file is elected, enable the OK Command button
    cmdOpen.Enabled = True
End Sub

Private Sub File1_DblClick()

    ' Populate the Login (frmLogon) SRCSAFE.INI text box
    ' with the selected file and then close this form
    Call PopulateSrcSafeini
    
End Sub

Private Sub Form_Load()

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub

Private Sub TEXT2_Change()
    File1.Pattern = Text2.Text
End Sub

' Populates the Srcsafe.ini information into the Logon Form
Public Sub PopulateSrcSafeini()
    If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
        frmLogon.txtSrcsafeini.Text = UCase(Dir1.Path + File1.FileName)
    Else
        frmLogon.txtSrcsafeini.Text = UCase(Dir1.Path + "\" + File1.FileName)
    End If
    Unload Me
End Sub
