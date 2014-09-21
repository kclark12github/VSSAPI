VERSION 5.00
Begin VB.Form frmLabel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label"
   ClientHeight    =   2985
   ClientLeft      =   4275
   ClientTop       =   6780
   ClientWidth     =   5235
   ClipControls    =   0   'False
   Icon            =   "label.frx":0000
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   5235
   Begin VB.TextBox txtComment 
      Height          =   1515
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   5070
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3975
      TabIndex        =   3
      Top             =   615
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3975
      TabIndex        =   2
      Top             =   135
      Width           =   1155
   End
   Begin VB.TextBox txtLabel 
      Height          =   285
      Left            =   585
      TabIndex        =   0
      Top             =   615
      Width           =   3255
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1110
      Width           =   705
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Label:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   660
      Width           =   435
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "Item:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   345
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim Response As Long
    
    ' Set On Error routine
    On Error GoTo Errhandler

    ' Label the item
    objVSSProject.Label Label:=txtLabel.Text, Comment:=txtComment.Text

    ' Check for errors
    If Err <> 0 Then
Errhandler:

        Response = MsgBox("Unable to label item '" + objVSSProject.Name + "'." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    
    ' Close Form
    Unload Me
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
