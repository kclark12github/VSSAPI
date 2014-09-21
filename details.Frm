VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Details"
   ClientHeight    =   5160
   ClientLeft      =   7050
   ClientTop       =   2235
   ClientWidth     =   5190
   ClipControls    =   0   'False
   Icon            =   "details.frx":0000
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5160
   ScaleWidth      =   5190
   Begin VB.TextBox txtAction 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   105
      TabIndex        =   11
      Top             =   2000
      Width           =   5010
   End
   Begin VB.TextBox txtLabelComment 
      Height          =   1020
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4050
      Width           =   5010
   End
   Begin VB.TextBox txtLabel 
      Height          =   285
      Left            =   735
      TabIndex        =   0
      Top             =   1450
      Width           =   4365
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4095
      TabIndex        =   3
      Top             =   180
      Width           =   960
   End
   Begin VB.TextBox txtComment 
      Height          =   1020
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2625
      Width           =   5010
   End
   Begin VB.Label lblLabelComment 
      AutoSize        =   -1  'True
      Caption         =   "Label Comment:"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   3780
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Label:"
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "User:"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1150
      Width           =   375
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   800
      Width           =   570
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   450
      Width           =   390
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   100
      Width           =   285
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   2370
      Width           =   705
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Dim RetVal As Long

    ' Check for changes and propmpt for save if found
    If objVSSVersion.Comment <> txtComment.Text Or objVSSVersion.Label <> txtLabel.Text Or objVSSVersion.LabelComment <> txtLabelComment.Text Then
        RetVal = MsgBox("Sorry. At this time the comment, label and label comment properties are read-only. Dang!", vbExclamation, AppTitle)
    End If
    
    ' Close the Form
    Unload Me

End Sub

Private Sub Form_Load()

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub
