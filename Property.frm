VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   5430
   ClientLeft      =   3630
   ClientTop       =   2445
   ClientWidth     =   5595
   ClipControls    =   0   'False
   Icon            =   "Property.frx":0000
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   4470
      TabIndex        =   6
      Top             =   5010
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4725
      Left            =   105
      TabIndex        =   7
      Top             =   180
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   8334
      _Version        =   327680
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Property.frx":014A
      Tab(0).ControlCount=   12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVersionComment"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLastLabel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblContains"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProjects"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFiles"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbFileType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtComment"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      TabCaption(1)   =   "Check Out Status"
      TabPicture(1)   =   "Property.frx":0166
      Tab(1).ControlCount=   13
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPurge"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "cmdRecover"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "lstDeleted"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "txtCheckOutComment"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "CheckedOutList"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblComment"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblCheckOutProject"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblCheckOutFolder"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblCheckOutSystem"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblCheckOutVersion"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblCheckOutDate"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblBy"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblFileNotCheckedOut"
      Tab(1).Control(12).Enabled=   0   'False
      TabCaption(2)   =   "Links"
      TabPicture(2)   =   "Property.frx":0182
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstLinks"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "lblLinksFor"
      Tab(2).Control(1).Enabled=   0   'False
      Begin VB.ListBox lstLinks 
         Height          =   3375
         Left            =   -74760
         MultiSelect     =   2  'Extended
         TabIndex        =   30
         Top             =   915
         Width           =   4875
      End
      Begin VB.CommandButton cmdPurge 
         Caption         =   "&Purge"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -70920
         TabIndex        =   5
         Top             =   1290
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdRecover 
         Caption         =   "Re&cover"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -70920
         TabIndex        =   4
         Top             =   810
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ListBox lstDeleted 
         Height          =   3765
         Left            =   -74805
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   750
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.TextBox txtCheckOutComment 
         Height          =   1065
         Left            =   -74805
         TabIndex        =   3
         Top             =   3510
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Frame Frame2 
         Height          =   780
         Left            =   2685
         TabIndex        =   17
         Top             =   2370
         Width           =   2490
         Begin VB.Label lblLabelVersion 
            AutoSize        =   -1  'True
            Caption         =   "Version:"
            Height          =   195
            Left            =   135
            TabIndex        =   19
            Top             =   195
            Width           =   570
         End
         Begin VB.Label lblLabelDate 
            AutoSize        =   -1  'True
            Caption         =   "Date:"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   480
            Width           =   390
         End
      End
      Begin VB.TextBox txtComment 
         Height          =   855
         Left            =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3600
         Width           =   4950
      End
      Begin VB.Frame Frame1 
         Height          =   780
         Left            =   150
         TabIndex        =   11
         Top             =   2370
         Width           =   2490
         Begin VB.Label lblVersionDate 
            AutoSize        =   -1  'True
            Caption         =   "Date:"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "Version:"
            Height          =   195
            Left            =   135
            TabIndex        =   12
            Top             =   195
            Width           =   570
         End
      End
      Begin VB.ComboBox cmbFileType 
         Height          =   315
         Left            =   750
         TabIndex        =   0
         Top             =   945
         Width           =   1110
      End
      Begin ComctlLib.ListView CheckedOutList 
         Height          =   3780
         Left            =   -74820
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   6668
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "User"
            Object.Tag             =   ""
            Text            =   "User"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   "CheckOutFolder"
            Object.Tag             =   ""
            Text            =   "Check Out Folder"
            Object.Width           =   4233
         EndProperty
      End
      Begin VB.Label lblLinksFor 
         AutoSize        =   -1  'True
         Caption         =   "Links for"
         Height          =   195
         Left            =   -74760
         TabIndex        =   31
         Top             =   555
         Width           =   600
      End
      Begin VB.Label lblFiles 
         AutoSize        =   -1  'True
         Caption         =   "Files:"
         Height          =   195
         Left            =   825
         TabIndex        =   29
         Top             =   1890
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblProjects 
         AutoSize        =   -1  'True
         Caption         =   "Projects"
         Height          =   195
         Left            =   825
         TabIndex        =   28
         Top             =   1635
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblContains 
         AutoSize        =   -1  'True
         Caption         =   "Contains:"
         Height          =   195
         Left            =   200
         TabIndex        =   27
         Top             =   1380
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   26
         Top             =   3285
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblCheckOutProject 
         AutoSize        =   -1  'True
         Caption         =   "Project:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   25
         Top             =   2900
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblCheckOutFolder 
         AutoSize        =   -1  'True
         Caption         =   "Folder"
         Height          =   195
         Left            =   -74805
         TabIndex        =   24
         Top             =   2505
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCheckOutSystem 
         AutoSize        =   -1  'True
         Caption         =   "Computer:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   23
         Top             =   2100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblCheckOutVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version:"
         Height          =   195
         Left            =   -74800
         TabIndex        =   22
         Top             =   1700
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblCheckOutDate 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   -74800
         TabIndex        =   21
         Top             =   1300
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblBy 
         AutoSize        =   -1  'True
         Caption         =   "By:"
         Height          =   195
         Left            =   -74800
         TabIndex        =   20
         Top             =   900
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblLastLabel 
         AutoSize        =   -1  'True
         Caption         =   "Last label:"
         Height          =   195
         Left            =   2745
         TabIndex        =   16
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label lblFileNotCheckedOut 
         AutoSize        =   -1  'True
         Caption         =   "File is not checked out"
         Height          =   195
         Left            =   -74805
         TabIndex        =   15
         Top             =   495
         Width           =   1605
      End
      Begin VB.Label lblVersionComment 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         Height          =   195
         Left            =   200
         TabIndex        =   14
         Top             =   3315
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Latest:"
         Height          =   195
         Left            =   200
         TabIndex        =   10
         Top             =   2175
         Width           =   480
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   200
         TabIndex        =   9
         Top             =   990
         Width           =   405
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   200
         TabIndex        =   8
         Top             =   615
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileType As Long
Dim objVSSPropertyObject As VSSItem

Private Sub cmdClose_Click()

    ' Close the properties form and set filetype
    If objVSSPropertyObject.Type = VSSITEM_FILE Then
        If frmProperties.cmbFileType.ListIndex = 1 Then
            objVSSPropertyObject.Binary = True
        Else
            objVSSPropertyObject.Binary = False
        End If
    End If
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdRecover_Click()
    
    ' User wants to recover a deleted item
    If cmdRecover.Caption = "Re&cover" Then
    
        Dim Count As Integer
        Dim Nodex As Node
        
        ' Set On Error Routine
        On Error Resume Next
        
        For Count = 0 To lstDeleted.ListCount - 1
            If lstDeleted.Selected(Count) Then
                lstDeleted.ListIndex = Count
                Set objVSSPropertyObject = objVSSDatabase.VSSItem(lstDeleted.Text, True)
                objVSSPropertyObject.Deleted = False
                
                ' Recover a file
                If objVSSPropertyObject.Type = VSSITEM_FILE Then
                   For Each objVSSVersion In objVSSPropertyObject.Versions
                        If objVSSVersion.VersionNumber = objVSSPropertyObject.VersionNumber Then
                            FileDate = objVSSVersion.Date
                            Exit For
                        End If
                    Next
                
                ' Recover a project
                Else
                    If frmMain.TreeView1.SelectedItem.Key = "$" Then
                        Set Nodex = frmMain.TreeView1.Nodes.Add("$", tvwChild, lstDeleted.Text, objVSSPropertyObject.Name, "Closed")
                        frmMain.TreeView1.Nodes.Item(frmMain.TreeView1.SelectedItem.Key).Sorted = True
                    Else
                        Set Nodex = frmMain.TreeView1.Nodes.Add(frmMain.TreeView1.SelectedItem.Key, tvwChild, lstDeleted.Text, objVSSPropertyObject.Name, "Closed")
                        frmMain.TreeView1.Nodes.Item(frmMain.TreeView1.SelectedItem.Key).Sorted = True
                    End If
                End If
            End If
        Next
        
        ' Repopulate lists
        Call frmMain.CheckForDeletedItems(objVSSPropertyObject.Parent)
        frmMain.RefreshFileList
        lstDeleted.Clear
        
        ' Disable command buttons
        cmdRecover.Enabled = False
        cmdPurge.Enabled = False
    
    ' Show details of selected checkout
    Else
    
        Dim objVSSCheckout As VSSCheckout
        Dim objUserListItem As ListItem
        Set objUserListItem = CheckedOutList.SelectedItem
        
        ' Populate the Check Out Details Form
        frmCheckoutDetails.lblFileName.Caption = "Name: " + objVSSObject.Name
        For Each objVSSCheckout In objVSSObject.Checkouts
            If objVSSCheckout.UserName = CheckedOutList.SelectedItem.Text And objVSSCheckout.LocalSpec = objUserListItem.SubItems(1) Then
                frmCheckoutDetails.lblBy.Caption = "By: " + objVSSCheckout.UserName
                frmCheckoutDetails.lblCheckOutDate.Caption = "Date: " + Str(objVSSCheckout.Date)
                frmCheckoutDetails.lblCheckOutVersion.Caption = "Version: " + Str(objVSSCheckout.VersionNumber)
                frmCheckoutDetails.lblCheckOutSystem.Caption = "Computer: " + objVSSCheckout.Machine
                frmCheckoutDetails.lblCheckOutFolder.Caption = "Folder: " + objVSSCheckout.LocalSpec
                frmCheckoutDetails.lblCheckOutProject.Caption = "Project: " + objVSSCheckout.Project
                frmCheckoutDetails.txtCheckOutComment.Text = objVSSCheckout.Comment
                Exit For
            End If
        Next
        
        ' Show the Check Out Details form
        frmCheckoutDetails.Show 1
        
    End If
End Sub

Private Sub cmdPurge_Click()
    
    Dim Count As Integer
    Dim Nodex As Node
    Dim RetVal As Long
    
    ' Destroy object
    For Count = 0 To lstDeleted.ListCount - 1
        If lstDeleted.Selected(Count) Then
            lstDeleted.ListIndex = Count
            Set objVSSPropertyObject = objVSSDatabase.VSSItem(lstDeleted.Text, True)
            If WarnPurge Then
                RetVal = MsgBox("Destroy cannot be undone; information will be lost permanently!" + vbCrLf + "Purge " + objVSSPropertyObject.Name + " anyway?", vbQuestion + vbYesNo, AppTitle)
            Else
                RetVal = vbYes
            End If
            If RetVal = vbYes Then
                objVSSPropertyObject.Destroy
            End If
        End If
    Next
    
    ' Repopulate list
    lstDeleted.Clear
    Call frmMain.CheckForDeletedItems(objVSSPropertyObject.Parent)
    
    ' Disable command buttons
    cmdRecover.Enabled = False
    cmdPurge.Enabled = False
End Sub

Private Sub Form_Load()

    ' Center Screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Populate Properties dialog
    cmbFileType.AddItem "Text"
    cmbFileType.AddItem "Binary"
    SSTab1.Tab = 0
    
    ' Instantiate the selected item
    If Selected = "Treeview1" Then
        
        ' Check if user has selected the root
        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            Set objVSSPropertyObject = objVSSDatabase.VSSItem("$/", False)
        Else
            Set objVSSPropertyObject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
        End If
    Else
        
        ' Set VSS Item to selected file in ListView control
        Set objVSSPropertyObject = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
    End If
    
    If objVSSPropertyObject.Type = VSSITEM_FILE Then
        If objVSSPropertyObject.Binary = True Then
            cmbFileType.ListIndex = 1
        Else
            cmbFileType.ListIndex = 0
        End If
    End If
End Sub

Private Sub lstDeleted_Click()
    
    ' Enable Recover and Purge Command Buttons
    cmdRecover.Enabled = True
    cmdPurge.Enabled = True
End Sub


