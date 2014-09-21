VERSION 5.00
Begin VB.Form frmWorkingFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Working Folder"
   ClientHeight    =   3060
   ClientLeft      =   4650
   ClientTop       =   5310
   ClientWidth     =   5565
   Icon            =   "working.frx":0000
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   5565
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4425
      TabIndex        =   5
      Top             =   960
      Width           =   1020
   End
   Begin VB.TextBox txtWorkingFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   3660
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4425
      TabIndex        =   3
      Top             =   90
      Width           =   1020
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Height          =   345
      Left            =   4425
      TabIndex        =   4
      Top             =   525
      Width           =   1020
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   3645
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   3645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folders:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   735
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Foldername:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Drives:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2355
      Width           =   495
   End
End
Attribute VB_Name = "frmWorkingFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreate_Click()

    Dim Response As Long

    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Create the directory if it doesn't exist
    If Dir(txtWorkingFolder, vbDirectory) = "" Then
        MkDir Trim(txtWorkingFolder)
    Else
        MsgBox ("The directory '" + txtWorkingFolder + "' already exists.")
    End If

    ' Check for errors
     If Err <> 0 Then
     
ErrHandler:
        
        Response = MsgBox("Unable to create folder." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
    
End Sub

Private Sub cmdOK_Click()

    Dim Response As Long
    
    ' Set On Error routine
    On Error Resume Next
    
    ' Create the directory if it doesn't exist
    If Dir(txtWorkingFolder, vbDirectory) = "" Then
        
        Response = MsgBox("The folder '" + txtWorkingFolder + "' doesn't exist, create?", vbYesNo, AppTitle)
        If Response = vbYes Then
            MkDir Trim(txtWorkingFolder.Text)
        End If
    Else
        Response = vbYes
    End If
    
    ' Set the Working Folder
    If Response = vbYes Then objVSSProject.LocalSpec = Trim(txtWorkingFolder.Text)
    
    ' Check for errors
    If Err <> 0 Then
        Response = MsgBox("Error setting working folder!" + vbCrLf + Err.Description + ".", vbExclamation, AppTitle)
        Err.Clear
    Else
        Call DisplayWorkingFolder(objVSSProject)
    End If
    
    ' Close the form
    If Response <> vbNo Then Unload Me
End Sub

Private Sub cmdCancel_Click()
    
    ' Close the form
    Unload Me
    
End Sub

Private Sub Dir1_Change()
    txtWorkingFolder.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    
    ' Center Screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Populate WorkingFolder Text Box
    txtWorkingFolder.Text = Dir1.Path
    
End Sub
