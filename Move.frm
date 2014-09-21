VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move "
   ClientHeight    =   3570
   ClientLeft      =   5895
   ClientTop       =   6690
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "Move.frx":0000
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   5730
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   4485
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   4485
      TabIndex        =   1
      Top             =   315
      Width           =   1035
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3075
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   4155
      _ExtentX        =   7329
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
   Begin VB.Label lblMoveTo 
      AutoSize        =   -1  'True
      Caption         =   "Move to:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   630
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    
    Dim objMoveItem As VSSItem
    Dim objMoveTo As VSSItem
    Dim Response As Long
    Dim Nodex As Node
    Dim ProjectToMoveTo As String
    Dim ProjectToMoveToPath As String
    
    ' Set On Error routine
    On Error GoTo Errhandler
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Instantiate project to move to
    If TreeView1.SelectedItem.Key <> "$" Then
        Set objMoveTo = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
    Else
        Set objMoveTo = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
    End If
    
    ' Moving a file
    If Selected = "Listview1" Then
        
        'Instantiate the file
        Set objMoveItem = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
        
        ' Move the file (not supported but the code is here to test the error message)
        objMoveItem.Move objMoveTo
        
    ' Moving a project
    Else
        
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If frmMain.TreeView1.SelectedItem.Key <> "$" Then
            Set objMoveItem = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
        Else
            Set objMoveItem = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Move the project
        objMoveItem.Move objMoveTo
        
        ' Clear the GUI
        frmMain.TreeView1.Nodes.Clear
        frmMain.ListView1.ListItems.Clear
        
        ' Repopulate GUI
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    
        ' Add Root Project to Treeview control
        frmMain.TreeView1.LineStyle = tvwRootLines
        Set Nodex = frmMain.TreeView1.Nodes.Add(, , "$", "$/", "Open")
        frmMain.TreeView1.Nodes(1).Selected = True
        frmMain.TreeView1.Nodes(1).Expanded = True
        
        ' Populate the Project and File List
        Call frmMain.PopulateMain(objVSSProject)
        frmMain.TreeView1.Nodes(1).Selected = True
        
        ' Put focus on the moved project
        If Right(objMoveTo.Spec, 1) = "/" Then
            ProjectToMoveToPath = objMoveTo.Spec + objMoveItem.Name
        Else
            ProjectToMoveToPath = objMoveTo.Spec + "/" + objMoveItem.Name
        End If
        If InStr(3, ProjectToMoveToPath, "/") <> 0 Then
            ProjectToMoveTo = Left(ProjectToMoveToPath, InStr(3, ProjectToMoveToPath, "/") - 1)
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
        End If
        While InStr(Len(ProjectToMoveTo) + 2, ProjectToMoveToPath, "/") <> 0
            ProjectToMoveTo = Left(ProjectToMoveToPath, InStr(Len(ProjectToMoveTo) + 2, ProjectToMoveToPath, "/") - 1)
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
        Wend
        
        ProjectToMoveTo = ProjectToMoveToPath
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
    End If
    
    ' Check for errors
    If Err <> 0 Then
Errhandler:
    
        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then
            Response = MsgBox("Unable to move item." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    
    ' Close form
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim Nodex As Node
    
    ' Intitlaize variables
    ItemCount = 0
    TreeView1.ImageList = frmMain.ImageList1
    lblMoveTo.Caption = "Move to $/"
    
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
        lblMoveTo.Caption = "Move to: " + objVSSProject.Spec
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        lblMoveTo.Caption = "Move to: $/"
    End If
    Call PopulateMain(objVSSProject)
    
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
    On Error GoTo Errhandler
    
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
Errhandler:

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
