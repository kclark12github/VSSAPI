VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "AddUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRights 
      Caption         =   "&Read Only"
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   1260
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2033
      TabIndex        =   4
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   503
      TabIndex        =   3
      Top             =   1575
      Width           =   1185
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   780
      Width           =   2505
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   1080
      TabIndex        =   0
      Top             =   270
      Width           =   2490
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   825
      Width           =   735
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   315
      Width           =   465
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rights As Boolean

Private Sub chkRights_Click()
    
    If chkRights.Value = 1 Then
        Rights = True
    Else
        Rights = False
    End If
    
End Sub

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim objUserToAdd As VSSUser
    Dim Response As Long
    Dim objUserListItem As ListItem
    
    ' Set On Error routine
    On Error GoTo Errhandler
    
    ' Add user
    Set objUserToAdd = objVSSDatabase.AddUser(txtName, txtPassword, Rights)
    
    ' Check for errors
    If Err <> 0 Then
Errhandler:
    
        Response = MsgBox("Unable to add user." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    Else
        Set objUserListItem = frmOptions.ListView1.ListItems.Add(, objUserToAdd.Name, objUserToAdd.Name)
        If chkRights.Value = 1 Then
            UserRights = "Read-Only"
        Else
            UserRights = "Read-Write"
        End If
        
        objUserListItem.SubItems(1) = UserRights
    End If
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize variables
    Rights = False
    chkRights.Value = 0
    
End Sub
