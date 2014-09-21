VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "ChangePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVerifyPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1425
      Width           =   2010
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1020
      Width           =   2010
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2010
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3510
      TabIndex        =   4
      Top             =   615
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3510
      TabIndex        =   3
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label lblVerify 
      AutoSize        =   -1  'True
      Caption         =   "Verify:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label lblNewPassword 
      AutoSize        =   -1  'True
      Caption         =   "New password:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label lblOldPassword 
      AutoSize        =   -1  'True
      Caption         =   "Old password:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   255
      Width           =   1005
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objVSSUser As VSSUser

Private Sub cmdCancel_Click()

    ' Close form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim Response As Long

    ' Set On Error routine
    On Error GoTo Errhandler:
    
    ' Instantiate the user object
    Set objVSSUser = objVSSDatabase.User(UserName)

    ' Verify entries and change password if entries are correct
    If txtOldPassword.Text = Password Then
        If txtNewPassword.Text = txtVerifyPassword.Text Then
            objVSSUser.Password = txtNewPassword.Text
            Response = MsgBox("Your password has been changed.", vbInformation, AppTitle)
            
            ' Close form
            Unload Me
            Exit Sub
        Else
            Response = MsgBox("Unable to verify new password.", vbExclamation, AppTitle)
        End If
    Else
        Response = MsgBox("Old password is incorrect.", vbExclamation, AppTitle)
    End If
    

    ' Check for errors
    If Err <> 0 Then
Errhandler:
    
        Response = MsgBox("Unable to change password." + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If

End Sub

Private Sub Form_Load()
    
    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub


