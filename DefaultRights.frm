VERSION 5.00
Begin VB.Form frmDefaultRights 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Default Rights"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "DefaultRights.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Default rights"
      Height          =   1725
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   3795
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "&Read Only"
         Height          =   270
         Left            =   255
         TabIndex        =   6
         Top             =   300
         Width           =   2670
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "&Check Out/Check In"
         Height          =   270
         Left            =   255
         TabIndex        =   5
         Top             =   630
         Width           =   2670
      End
      Begin VB.CheckBox chkAddRenameDelete 
         Caption         =   "&Add/Rename/Delete"
         Height          =   270
         Left            =   255
         TabIndex        =   4
         Top             =   975
         Width           =   2670
      End
      Begin VB.CheckBox chkDestroy 
         Caption         =   "&Destroy"
         Height          =   270
         Left            =   255
         TabIndex        =   3
         Top             =   1305
         Width           =   2670
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   668
      TabIndex        =   0
      Top             =   2010
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2198
      TabIndex        =   1
      Top             =   2010
      Width           =   1185
   End
End
Attribute VB_Name = "frmDefaultRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDestroy_Click()
    
    If chkDestroy.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 1
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkReadOnly.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkReadOnly.Value = 1
    End If
    
End Sub

' Set Default Rights to Read Only

Private Sub chkReadOnly_Click()
    
    If chkReadOnly.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkReadOnly.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkReadOnly.Value = 0
    End If
    
End Sub

Private Sub chkUpdate_Click()

    If chkUpdate.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 1
        chkReadOnly.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkReadOnly.Value = 1
    End If
    
End Sub

Private Sub chkAddRenameDelete_Click()

    If chkAddRenameDelete.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkReadOnly.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 1
        chkReadOnly.Value = 1
    End If

End Sub
    
Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Dim DefaultRights As Long
    Dim Response As Long
    
    ' Set On Error routine
    On Error GoTo Errhandler
    
    ' Set Default Project Rights
    If chkReadOnly.Value = 0 Then
    
        DefaultRights = 0
    
    ElseIf chkReadOnly.Value = 1 And chkUpdate.Value = 0 Then
    
        DefaultRights = VSSRIGHTS_READ
        
    ElseIf chkUpdate.Value = 1 And chkAddRenameDelete.Value = 0 Then
        
            DefaultRights = VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
            
    ElseIf chkAddRenameDelete.Value = 1 And chkDestroy.Value = 0 Then
        
            DefaultRights = VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
            
    ElseIf chkDestroy.Value = 1 Then
        
            DefaultRights = VSSRIGHTS_ALL
    
    End If
    
    ' Set default rights
    objVSSDatabase.DefaultProjectRights = DefaultRights
    
    ' Check for errors
    If Err <> 0 Then
Errhandler:
    
        Response = MsgBox("Unable to set default rights." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If

    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize the form
    Select Case objVSSDatabase.DefaultProjectRights
    
        Case VSSRIGHTS_READ
            
            chkReadOnly.Value = 1
        
        Case VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
        
            chkReadOnly.Value = 1
            chkUpdate.Value = 1
            
        Case VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
        
            chkReadOnly.Value = 1
            chkUpdate.Value = 1
            chkAddRenameDelete.Value = 1
            
        Case VSSRIGHTS_DESTROY + VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
        
            chkReadOnly.Value = 1
            chkUpdate.Value = 1
            chkAddRenameDelete.Value = 1
            chkDestroy.Value = 1
        
    End Select
    
End Sub
