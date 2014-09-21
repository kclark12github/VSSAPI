VERSION 5.00
Begin VB.Form frmCheckoutDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Out Status"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "CheckoutDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCheckOutComment 
      Height          =   1065
      Left            =   60
      TabIndex        =   1
      Top             =   3090
      Width           =   4995
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   3945
      TabIndex        =   0
      Top             =   165
      Width           =   1050
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   75
      Width           =   465
   End
   Begin VB.Label lblBy 
      AutoSize        =   -1  'True
      Caption         =   "By:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   225
   End
   Begin VB.Label lblCheckOutDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   885
      Width           =   390
   End
   Begin VB.Label lblCheckOutVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1275
      Width           =   570
   End
   Begin VB.Label lblCheckOutSystem 
      AutoSize        =   -1  'True
      Caption         =   "Computer:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lblCheckOutFolder 
      AutoSize        =   -1  'True
      Caption         =   "Folder"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2085
      Width           =   435
   End
   Begin VB.Label lblCheckOutProject 
      AutoSize        =   -1  'True
      Caption         =   "Project:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2475
      Width           =   540
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2865
      Width           =   705
   End
End
Attribute VB_Name = "frmCheckoutDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub
