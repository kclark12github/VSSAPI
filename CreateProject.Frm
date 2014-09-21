VERSION 5.00
Begin VB.Form frmCreateProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Project in"
   ClientHeight    =   2685
   ClientLeft      =   4665
   ClientTop       =   4815
   ClientWidth     =   5385
   ClipControls    =   0   'False
   Icon            =   "CreateProject.frx":0000
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComment 
      Height          =   1515
      Left            =   105
      TabIndex        =   1
      Top             =   1080
      Width           =   5115
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4170
      TabIndex        =   3
      Top             =   600
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   4170
      TabIndex        =   2
      Top             =   165
      Width           =   1050
   End
   Begin VB.TextBox txtProjectName 
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   150
      Width           =   3120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "frmCreateProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    
    ' Close Form
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim Nodex As Node
    Dim Response As Long
    
    ' Set on error routine
    On Error Resume Next
    
    ' Refresh TreeView control to make sure the project being
    ' added does not exist already
    frmMain.TreeView1.Refresh
    
    ' Attempt to add project to database
    objVSSProject.NewSubproject Name:=txtProjectName.Text, Comment:=txtComment.Text
    
    ' Add was successful
    If Err = 0 Then
        
        ' Add new project to Project List
        Set Nodex = frmMain.TreeView1.Nodes.Add(frmMain.TreeView1.SelectedItem.Key, tvwChild, frmMain.TreeView1.SelectedItem.Key + "/" + txtProjectName.Text, txtProjectName.Text, "Closed")
        
        ' Sort TreeView control
        frmMain.TreeView1.Nodes.Item(frmMain.TreeView1.SelectedItem.Key).Sorted = True
        
        ' Close the Create Project dialog
        Unload Me
    
    ' Otherwise check for errors and respond as appropriate
    ElseIf Err = -2147166572 Then
        Response = MsgBox("Sorry, a project of the name " + txtProjectName.Text + " already exists in this project.", vbExclamation, AppTitle)
    Else
        Response = MsgBox("Unable to add project " + txtProjectName.Text + "." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    End If
End Sub

Private Sub Form_Load()

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub txtProjectName_Change()
    
    ' Enable OK button if New Project name
    ' contains text
    If txtProjectName.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
