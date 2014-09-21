VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmShare 
   Caption         =   "Share with"
   ClientHeight    =   4185
   ClientLeft      =   5910
   ClientTop       =   6705
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "share.frx":0000
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   5730
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive (projects only)"
      Height          =   225
      Left            =   2880
      TabIndex        =   3
      Top             =   3915
      Width           =   2565
   End
   Begin VB.CheckBox chkBranch 
      Caption         =   "&Branch after share (files only)"
      Enabled         =   0   'False
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   3915
      Width           =   2565
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Vie&w"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4500
      TabIndex        =   6
      Top             =   1080
      Width           =   1035
   End
   Begin VB.CommandButton cmdShare 
      Caption         =   "&Share"
      Default         =   -1  'True
      Height          =   330
      Left            =   4500
      TabIndex        =   5
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   4500
      TabIndex        =   4
      Top             =   135
      Width           =   1035
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5900
      View            =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "NAME"
         Object.Tag             =   "NAME"
         Text            =   "Name:"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3075
      Left            =   2040
      TabIndex        =   1
      Top             =   765
      Width           =   2280
      _ExtentX        =   4022
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
   Begin VB.Label lblContents 
      Caption         =   "ProjectName"
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   495
      Width           =   2250
   End
   Begin VB.Label lblFileToShare 
      AutoSize        =   -1  'True
      Caption         =   "File to share:"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   195
      Width           =   900
   End
   Begin VB.Label lblProjects 
      AutoSize        =   -1  'True
      Caption         =   "Projects:"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShareFlag As Long

Private Sub chkRecursive_Click()
    
    ' Set share flag
    If chkRecursive.Value = 1 Then
        ShareFlag = VSSFLAG_RECURSYES
    Else
        ShareFlag = VSSFLAG_RECURSNO
    End If
    
End Sub

Private Sub cmdClose_Click()
    
    ' Close Form
    Unload Me
    
End Sub

' Share the File(s)

Private Sub cmdShare_Click()
    
    Dim Response As Long
    Dim objListItem As ListItem
    Dim objFileListItem As ListItem
    Dim objVSSFile As VSSItem
    Dim FileToBranch As String
    Dim ProjectToBranch As String
    Dim objFiletoBranch As VSSItem
    Dim objProjecttoBranch As VSSItem
    Dim Count As Integer
    Dim SelectedCount As Integer
    Dim ObjectType As String
    Dim objVSSProjectToShare As VSSItem
    Dim Nodex As Node
                    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    SelectedCount = 0
    
    ' Instantiate selected project item
    If frmMain.TreeView1.SelectedItem.Key = "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
    End If
    
    ' Iterate through the file list
    For Count = 1 To ListView1.ListItems.Count
        
        Set objFileListItem = ListView1.ListItems(Count)
    
        ' File is selected
        If objFileListItem.Selected = True Then
        
            ' Instantiate the file item
            Set objVSSFile = objVSSDatabase.VSSItem(objFileListItem.Key, False)
            SelectedCount = SelectedCount + 1
            ObjectType = "file"
            
            ' Share the file and update the main form's GUI
            objVSSProject.Share pIItem:=objVSSFile, Comment:="", iFlags:=0

            ' User wants to branch file
            If chkBranch.Value = 1 Then
                
                ' Wait for share to complete then branch
                FileToBranch = objVSSProject.Spec + objVSSFile.Name
                Set objFiletoBranch = objVSSDatabase.VSSItem(FileToBranch, False)
                objFiletoBranch.Branch Comment:="", iFlags:=0
                Call frmMain.AddFileItem(objFileListItem.Key, True, "Leaf")
            Else
                Call frmMain.AddFileItem(objFileListItem.Key, True, "Shared")
            End If
        End If
ShareNext:

    Next
            
    ' Sharing a project
    If SelectedCount = 0 Then
        ObjectType = "project"
        
        ' Instantiate the project object
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProjectToShare = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objVSSProjectToShare = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
       
        ' Share the file and update the main form's GUI
        objVSSProject.Share pIItem:=objVSSProjectToShare, Comment:="", iFlags:=ShareFlag
        
        ' Add project to Treeview control
        Set Nodex = frmMain.TreeView1.Nodes.Add(objVSSProject.Spec, tvwChild, objVSSProject.Spec + "/" + objVSSProjectToShare.Name, objVSSProjectToShare.Name, "Closed")
        Nodex.Parent.Sorted = True
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
            
        Response = MsgBox("Unable to Share " + ObjectType + "." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        If SelectedCount <> 0 Then Resume ShareNext
        Err.Clear
    End If
        
End Sub

Private Sub cmdView_Click()

    Call frmMain.ViewFile(ListView1.SelectedItem.Key)
    
End Sub

Private Sub Form_Load()

    Dim Nodex As Node
    
    ' Intitlaize variables
    ItemCount = 0
    TreeView1.ImageList = frmMain.ImageList1
    lblContents.Caption = "$/"
    ShareFlag = VSSFLAG_RECURSNO
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Create VSS Database object and set current item to $/ (root project)
    Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    
    ' Add Root Project to Treeview control
    TreeView1.LineStyle = tvwRootLines
    Set Nodex = TreeView1.Nodes.Add(, , "$", "$/", "Closed")
    TreeView1.Nodes(1).Selected = True
    
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

Private Sub Form_Resize()

    On Error Resume Next

    ' Don't you wish the real SourceSafe did this?
    cmdClose.Left = frmShare.Width - 1350
    cmdShare.Left = cmdClose.Left
    cmdView.Left = cmdClose.Left
    TreeView1.Width = cmdClose.Left - 180 - TreeView1.Left
    ListView1.Height = frmShare.Height - 1245
    TreeView1.Height = frmShare.Height - 1515
    chkBranch.Top = TreeView1.Top + TreeView1.Height + 75
    
End Sub

Private Sub ListView1_Click()
    
    chkBranch.Enabled = True
    cmdView.Enabled = True
    chkRecursive.Enabled = False
    
End Sub

Private Sub TreeView1_Click()
    
    Dim Nodex As Node
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set glyphs for all projects to the closed icon
    For Each Nodex In TreeView1.Nodes
        If Nodex.Image = "Open" Then Nodex.Image = "Closed"
    Next
    
    ' Set checkbox state
    chkBranch.Enabled = False
    chkRecursive.Enabled = True
    
    ' Set glyph for selected project to open icon
    TreeView1.SelectedItem.Image = "Open"
    
    ' Clear the ListView control of all files
    ListView1.ListItems.Clear
    
    ' Check if user has selected the root project and set the VSSItem as appropriate
    If TreeView1.SelectedItem.Key <> "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        lblContents.Caption = objVSSProject.Spec
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        lblContents.Caption = "$/"
    End If
    Call PopulateMain(objVSSProject)
    
    ' Set View Command Buttton's Enabled state
    cmdView.Enabled = False
    
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
    
    ' Clear the File List
    ListView1.ListItems.Clear

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
    Dim objFileListItem As ListItem
    
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
            
            ' Current Object is a file
            Case 1
                
                Set objFileListItem = ListView1.ListItems.Add(, TreeView1.SelectedItem.Key + "/" + objVSSObject.Name, objVSSObject.Name, , "Leaf")
                
                ' Tally number of items in project for display purposes
                ItemCount = ItemCount + 1
            
            ' Unknown object type
            Case Else
                MsgBox ("Unknown object type encountered during Node population!")
        End Select
    Next
    
End Sub
