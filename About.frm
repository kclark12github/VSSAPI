VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&SourceSafe Automation is cool!"
      Height          =   420
      Left            =   720
      TabIndex        =   0
      Top             =   1395
      Width           =   2655
   End
   Begin VB.Label lblAboutInfo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   960
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   3720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Set caption
    frmAbout.Caption = "About " + AppTitle
    lblAboutInfo.Caption = AppTitle + " Version 2.0" + vbCrLf + vbCrLf + "This sample was created by Tim Winter using Visual Basic 5.0. Please see the Readme file for additional information."

End Sub
