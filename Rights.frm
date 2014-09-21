VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmRights 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Project Rights"
   ClientHeight    =   3480
   ClientLeft      =   5895
   ClientTop       =   6690
   ClientWidth     =   6660
   ClipControls    =   0   'False
   Icon            =   "Rights.frx":0000
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   6660
   Begin VB.Frame Frame1 
      Caption         =   "User Rights"
      Height          =   2310
      Left            =   4020
      TabIndex        =   9
      Top             =   225
      Width           =   2550
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remo&ve"
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Removes the rights for the current project (this will cause the project to inherit rights from it's parent)"
         Top             =   1770
         Width           =   1035
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set"
         Height          =   330
         Left            =   195
         TabIndex        =   5
         ToolTipText     =   "Set the rights for the current project"
         Top             =   1770
         Width           =   1035
      End
      Begin VB.CheckBox chkDestroy 
         Caption         =   "&Destroy"
         Height          =   270
         Left            =   270
         TabIndex        =   4
         Top             =   1290
         Width           =   1905
      End
      Begin VB.CheckBox chkAddRenameDelete 
         Caption         =   "&Add/Rename/Delete"
         Height          =   270
         Left            =   270
         TabIndex        =   3
         Top             =   955
         Width           =   1905
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "&Check Out/Check In"
         Height          =   270
         Left            =   270
         TabIndex        =   2
         Top             =   620
         Width           =   1905
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "&Read"
         Height          =   270
         Left            =   270
         TabIndex        =   1
         Top             =   285
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Default         =   -1  'True
      Height          =   330
      Left            =   5535
      TabIndex        =   7
      Top             =   3060
      Width           =   1035
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3075
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   5424
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblProjects 
      AutoSize        =   -1  'True
      Caption         =   "Project list:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   765
   End
End
Attribute VB_Name = "frmRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentProject As String

Private Sub cmdClose_Click()
    
    ' Close the form
    Unload Me
    
End Sub

Private Sub cmdSet_Click()

    Dim Response As Long

    ' Set On Error Routine
    On Error GoTo ErrHandler

    ' Set the current Project Rights
    If chkRead.Value = 0 Then
        CurrentUser.ProjectRights(CurrentProject) = 0
    ElseIf chkRead.Value = 1 And chkUpdate.Value = 0 Then
        CurrentUser.ProjectRights(CurrentProject) = VSSRIGHTS_READ
    ElseIf chkUpdate.Value = 1 And chkAddRenameDelete.Value = 0 Then
        CurrentUser.ProjectRights(CurrentProject) = VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
    ElseIf chkAddRenameDelete.Value = 1 And chkDestroy.Value = 0 Then
        CurrentUser.ProjectRights(CurrentProject) = VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ
    ElseIf chkDestroy.Value = 1 Then
        CurrentUser.ProjectRights(CurrentProject) = VSSRIGHTS_ALL
    End If
    Call GetProjectRights(CurrentProject)
    
ErrHandler:
    If Err <> 0 Then
        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If

End Sub

Private Sub cmdRemove_Click()
    
    Dim Response As Long

    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Remove and "set" project rights. This will
    ' cause the project to inherit rights from
    ' it's parent
    CurrentUser.RemoveProjectRights (CurrentProject)
    Call GetProjectRights(CurrentProject)
    
ErrHandler:
    If Err <> 0 Then
        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If

End Sub

Private Sub Form_Load()

    Dim Nodex As Node
    
    ' Intitlaize variables
    ItemCount = 0
    TreeView1.ImageList = frmMain.ImageList1
    CurrentProject = "$/"
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Create VSS Database object and set current item to $/ (root project)
    Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    
    ' Add Root Project to Treeview control
    TreeView1.LineStyle = tvwRootLines
    Set Nodex = TreeView1.Nodes.Add(, , "$", "$/", "Open")
    TreeView1.Nodes(1).Selected = True
    TreeView1.Nodes(1).Expanded = True
    
    ' Populate the Project and File List
    Call PopulateMain(objVSSProject)
    
    ' Get Rights for Root Project
    Call GetProjectRights(CurrentProject)
    If CurrentUser.ReadOnly = True Then
        chkUpdate.Enabled = False
        chkAddRenameDelete.Enabled = False
        chkDestroy.Enabled = False
        cmdSet.Enabled = False
        cmdRemove.Enabled = False
    End If
    
End Sub

' This routine is passed a project item as a parameter. It checks for existing
' sub projects in the passed project and is used for populating the TreeView
' Control. If the passed project contains sub projects, it must be added to
' the control allowing the project to be "expanded"

Public Sub PopulateSubProjects(objVSSProject As VSSItem)

    Dim Nodex As Node
    
    ' Iterate through each item of the project (false means ignore deleted)
    For Each objVSSObject In objVSSProject.Items(False)
        
        ' If a sub project (type = 0) is found then add it as a child to the
        ' current project
        If objVSSObject.Type = 0 Then
            Set Nodex = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key + "/" + objVSSProject.Name, tvwChild, TreeView1.SelectedItem.Key + "/" + objVSSProject.Name + "/" + objVSSObject.Name, objVSSObject.Name, "Closed")
        End If
    Next
End Sub

Private Sub TreeView1_Click()
    
    Dim Nodex As Node
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set glyphs for all projects to the closed icon
    For Each Nodex In TreeView1.Nodes
        If Nodex.Image = "Open" Then Nodex.Image = "Closed"
    Next
    
    ' Set glyph for selected project to open icon
    TreeView1.SelectedItem.Image = "Open"
    
    ' Check if user has selected the root project and set the VSSItem as appropriate
    If TreeView1.SelectedItem.Key <> "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
    End If
    CurrentProject = objVSSProject.Spec
    Call PopulateMain(objVSSProject)
    GetProjectRights (objVSSProject.Spec)
    
    ' Reset mousepointer
    MousePointer = vbNormal
    
End Sub

Private Sub TreeView1_Collapse(ByVal Node As Node)

    Node.Image = "Closed"
    
End Sub

' This routine is called when the user expands a project in the
' TreeView control.

Private Sub TreeView1_Expand(ByVal Node As Node)

    Dim Response As Long
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Open the current Project glyph
    Node.Image = "Open"
    TreeView1.Nodes(Node.Index).Selected = True

    ' Check to see if we are expanding the Root Project and
    ' instantiate the VSSItem as appropriate
    Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
    
    ' Call routine to populate the TreeView and ListView controls
    Call PopulateMain(objVSSProject)
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:

        ' If error is from a non unique key we can ignore it
        If Err <> 35602 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        End If
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal

End Sub

' This routine populates the ListView and TreeView controls as appropriate

Public Sub PopulateMain(objVSSProject As Object)

    Dim Nodex As Node
    
    ' Set On Error routine
    On Error Resume Next
    
    ' Set Item Count
    ItemCount = 0
    
    ' Iterate through all items in current project (false means ignore deleted items)
    For Each objVSSObject In objVSSProject.Items(False)
    
        ' Check to see what type of object we have
        Select Case objVSSObject.Type
        
            ' Current item is a project
            Case 0
                
                ' Add project to Treeview control
                Set Nodex = TreeView1.Nodes.Add("$", tvwChild, TreeView1.SelectedItem.Key + "/" + objVSSObject.Name, objVSSObject.Name, "Closed")
                
                ' Call procedure to check for existing sub projects of this
                ' project (this poplates the control with + signs as needed)
                Call PopulateSubProjects(objVSSObject)
                
                If Err = 0 Then Err.Clear
            
        End Select
    Next
    
End Sub

' Set Default Rights to Read Only

Private Sub chkRead_Click()
    
    If chkRead.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkRead.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkRead.Value = 0
    End If
    
End Sub

Private Sub chkUpdate_Click()

    If chkUpdate.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 1
        chkRead.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 0
        chkRead.Value = 1
    End If
    
End Sub

Private Sub chkAddRenameDelete_Click()

    If chkAddRenameDelete.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkRead.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 0
        chkUpdate.Value = 1
        chkRead.Value = 1
    End If

End Sub
    
Private Sub chkDestroy_Click()
    
    If chkDestroy.Value = 1 Then
    
        ' Set checkboxes as appropriate
        chkDestroy.Value = 1
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkRead.Value = 1

    Else
        
        ' Set checkboxes as appropriate
        chkDestroy.Value = 0
        chkAddRenameDelete.Value = 1
        chkUpdate.Value = 1
        chkRead.Value = 1
    End If
    
End Sub

Public Sub GetProjectRights(ProjectName As String)

    Dim Response As Long

    ' Set On Error Routine
    On Error GoTo ErrHandler

    Select Case CurrentUser.ProjectRights(ProjectName)
        
        Case VSSRIGHTS_ALL, VSSRIGHTS_INHERITED + VSSRIGHTS_ALL
            
            chkDestroy.Value = 1
            chkDestroy.Value = 1
            chkDestroy_Click
        
        Case VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ, VSSRIGHTS_ADDRENREM + VSSRIGHTS_CHKUPD + VSSRIGHTS_READ + VSSRIGHTS_INHERITED
            
            chkAddRenameDelete.Value = 1
            chkAddRenameDelete.Value = 1
            chkAddRenameDelete_Click
        
        Case VSSRIGHTS_CHKUPD + VSSRIGHTS_READ, VSSRIGHTS_CHKUPD + VSSRIGHTS_READ + VSSRIGHTS_INHERITED
            
            chkUpdate.Value = 1
            chkUpdate.Value = 1
            chkUpdate_Click
        
        Case VSSRIGHTS_READ, VSSRIGHTS_READ + VSSRIGHTS_INHERITED
        
            chkRead.Value = 1
            chkRead.Value = 1
            chkRead_Click
    
    End Select
    
ErrHandler:
    If Err <> 0 Then
        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
    
    
End Sub
