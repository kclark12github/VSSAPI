VERSION 5.00
Begin VB.Form frmEditUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   ClipControls    =   0   'False
   Icon            =   "EditUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "&Read Only"
      Height          =   345
      Left            =   750
      TabIndex        =   1
      Top             =   720
      Width           =   1515
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   750
      TabIndex        =   0
      Top             =   225
      Width           =   2745
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1305
      TabIndex        =   2
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   300
      Width           =   465
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

    Dim Response As Long

    ' Set Error Routine
    On Error GoTo ErrHandler
    
    ' Automation returns an error (by design) if the user
    ' attempts to set the name property to it's current
    ' value, so change name only if a new name is entered
    If UCase(CurrentUser.Name) <> UCase(frmEditUser.txtName.Text) Then
        CurrentUser.Name = frmEditUser.txtName.Text
        frmOptions.ListView1.SelectedItem.Text = CurrentUser.Name
    End If
    If chkReadOnly.Value = 1 Then
        CurrentUser.ReadOnly = True
        frmOptions.ListView1.SelectedItem.SubItems(1) = "Read-Only"
    Else
        CurrentUser.ReadOnly = False
        frmOptions.ListView1.SelectedItem.SubItems(1) = "Read-Write"
    End If
    
    ' Close the form
    Unload Me
    
    ' Check for Errors
    If Err <> 0 Then
ErrHandler:

        Response = MsgBox("Unable to edit user." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
End Sub

Private Sub Form_Load()

    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub
