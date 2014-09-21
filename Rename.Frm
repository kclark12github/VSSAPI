VERSION 5.00
Begin VB.Form frmRename 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename"
   ClientHeight    =   1215
   ClientLeft      =   5085
   ClientTop       =   4680
   ClientWidth     =   4320
   ClipControls    =   0   'False
   Icon            =   "rename.frx":0000
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1215
   ScaleWidth      =   4320
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Top             =   105
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3120
      TabIndex        =   2
      Top             =   570
      Width           =   1050
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   885
      TabIndex        =   0
      Top             =   105
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rename:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim Response As Long
    Dim ItemKey As String
    Dim NodChild As Node
    Dim ChildName As String
    Dim NodeNumber As Integer
    Dim NodeToSort As Integer
    
    ' Prepare to reset item keys
    If objVSSProject.Type = VSSITEM_FILE Then

        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            ItemKey = "$/"
        Else
            ItemKey = frmMain.TreeView1.SelectedItem.Key + "/"
        End If
    Else
        
        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            ItemKey = "$/"
            NodeToSort = 1
        Else
            ItemKey = frmMain.TreeView1.SelectedItem.Parent.Key + "/"
            NodeToSort = frmMain.TreeView1.SelectedItem.Parent.Index
        End If

    End If
    
    ' Set On Error routine
    On Error GoTo Errhandler
    
    ' Rename the item
    objVSSProject.Name = Trim(txtNewName)
    
    ' Reset the item key
    If objVSSProject.Type = VSSITEM_FILE Then
        frmMain.ListView1.SelectedItem.Text = objVSSProject.Name
        frmMain.ListView1.SelectedItem.Key = ItemKey + objVSSProject.Name
        frmMain.ListView1.Refresh
    Else
        frmMain.TreeView1.SelectedItem.Text = objVSSProject.Name
        frmMain.TreeView1.SelectedItem.Key = ItemKey + objVSSProject.Name
        
        ' Reset Item Keys of any subprojects
        If frmMain.TreeView1.Nodes(frmMain.TreeView1.SelectedItem.Index).Children > 0 Then
            
            ' Get first child's text, and set NodeNumber to its index value.
            ChildName = frmMain.TreeView1.SelectedItem.Child.Text
            NodeNumber = frmMain.TreeView1.SelectedItem.Child.Index
    
            ' While N is not the index of the child node's
            ' last sibling, get next sibling's text.
            While NodeNumber <= frmMain.TreeView1.SelectedItem.Child.LastSibling.Index
    
                frmMain.TreeView1.Nodes(NodeNumber).Key = frmMain.TreeView1.Nodes(NodeNumber).Parent.Key + "/" + ChildName
                
                ' Reset NodeNumber to next sibling's index
                If NodeNumber < frmMain.TreeView1.SelectedItem.Child.LastSibling.Index Then
                    NodeNumber = frmMain.TreeView1.Nodes(NodeNumber).Next.Index
                    ChildName = frmMain.TreeView1.Nodes(NodeNumber).Text
                Else
                    NodeNumber = NodeNumber + 1
                End If
            Wend
        End If
    End If
    
    ' Refresh sort order
    If objVSSProject.Type = VSSITEM_FILE Then
        frmMain.ListView1.Sorted = True
    Else
        frmMain.TreeView1.Nodes(NodeToSort).Sorted = True
    End If
   
    ' Check for errors
    If Err <> 0 Then
Errhandler:

        Response = MsgBox("Unable to rename object!" + vbCrLf + Err.Description, vbExclamation, AppTitle)
    End If
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub txtNewName_Change()
    
    ' Enable\disable OK button as appropriate
    If txtNewName.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
