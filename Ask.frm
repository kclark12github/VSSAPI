VERSION 5.00
Begin VB.Form frmAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "Ask.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   5580
      TabIndex        =   2
      Top             =   210
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   5580
      TabIndex        =   1
      Top             =   660
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Do you want to"
      Height          =   1185
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5220
      Begin VB.OptionButton optEdit 
         Caption         =   "&Check out this file and edit it in your working folder"
         Height          =   360
         Left            =   270
         TabIndex        =   4
         Top             =   660
         Width           =   4215
      End
      Begin VB.OptionButton optView 
         Caption         =   "&View SourceSafe's copy of this file"
         Height          =   360
         Left            =   270
         TabIndex        =   3
         Top             =   285
         Value           =   -1  'True
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Dim WhatToDo As Integer
    
    If optView.Value = 0 Then
        WhatToDo = 0
    Else
        WhatToDo = 1
    End If
    
    ' Close Form
    Unload Me

    If WhatToDo = 0 Then
        frmMain.mnuEditFile_Click
    Else
        frmMain.mnuViewFile_Click
    End If

End Sub

Private Sub Form_Load()

   ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

End Sub
