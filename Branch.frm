VERSION 5.00
Begin VB.Form frmBranch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Branch"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Branch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   465
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Top             =   120
      Width           =   2940
   End
   Begin VB.TextBox txtComment 
      Height          =   1485
      Left            =   60
      TabIndex        =   2
      Top             =   1650
      Width           =   4530
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3540
      TabIndex        =   1
      Top             =   570
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3540
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1395
      Width           =   705
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

' Branch the file

Private Sub cmdOK_Click()
        
    Dim Count As Integer
    Dim Response As Long
    Dim objVSSFile As VSSItem
    Dim FoundSelectedItem As Boolean
    Dim SingleItem As Boolean
    
    ' Initialize variables
    FoundSelectedItem = False
    SingleItem = False
    
    ' Set On Error routine
    On Error GoTo Errhandler
    
    ' Hide the Form
    frmBranch.Hide
    
    ' Iterate through File List
    For Count = 1 To frmMain.ListView1.ListItems.Count
    
        ' File is selected
        If frmMain.ListView1.ListItems(Count).Selected = True Then
            FoundSelectedItem = True
            
            ' Instantiate VSSItem
            Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(Count).Key, False)
        
            ' Branch
            objVSSFile.Branch Comment:=txtComment.Text, iFlags:=0
        
        End If
ContinueAfterError:

    Next
    
    If Not FoundSelectedItem Then
        SingleItem = True
        Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
        
        ' Branch
        objVSSFile.Branch Comment:=txtComment.Text, iFlags:=0
    End If
    
    
    ' Check for errors
    If Err <> 0 Then
Errhandler:
        
        Response = MsgBox("Unable to Branch file '" + objVSSFile.Name + "'." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
        If Not SingleItem Then
            Resume ContinueAfterError
        End If
    End If
    
    ' Close Form
    Unload Me

End Sub

Private Sub Form_Load()

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub
