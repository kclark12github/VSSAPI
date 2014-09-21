VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmAddFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add File to "
   ClientHeight    =   3810
   ClientLeft      =   3135
   ClientTop       =   3900
   ClientWidth     =   8085
   ClipControls    =   0   'False
   Icon            =   "add.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3810
   ScaleWidth      =   8085
   Begin ComctlLib.ListView ListView1 
      Height          =   2265
      Left            =   60
      TabIndex        =   19
      Top             =   825
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkStoreLatestVersion 
      Caption         =   "Store only latest &version"
      Height          =   195
      Left            =   4980
      TabIndex        =   9
      Top             =   2745
      Width           =   3015
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "Recur&sive"
      Height          =   300
      Left            =   4980
      TabIndex        =   7
      Top             =   2115
      Width           =   2970
   End
   Begin VB.ComboBox cmbFileType 
      Height          =   315
      Left            =   2460
      TabIndex        =   5
      Top             =   3405
      Width           =   2595
   End
   Begin VB.CheckBox chkRemoveLocal 
      Caption         =   "&Remove local copy"
      Height          =   300
      Left            =   4980
      TabIndex        =   8
      Top             =   2400
      Width           =   2970
   End
   Begin VB.TextBox txtComment 
      Height          =   1200
      Left            =   4980
      TabIndex        =   6
      Top             =   855
      Width           =   3120
   End
   Begin VB.TextBox txtFilePattern 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "*.*"
      Top             =   3405
      Width           =   2235
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   7005
      TabIndex        =   11
      Top             =   465
      Width           =   1020
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   7020
      TabIndex        =   10
      Top             =   30
      Width           =   1020
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   2580
      Width           =   2220
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   2460
      TabIndex        =   3
      Top             =   855
      Width           =   2160
   End
   Begin VB.FileListBox lstFiles 
      Height          =   1260
      Left            =   345
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4110
      Width           =   7560
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   405
      Width           =   2235
   End
   Begin VB.Label lblFileType 
      Caption         =   "File Type:"
      Height          =   300
      Left            =   2460
      TabIndex        =   18
      Top             =   3135
      Width           =   2430
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   4980
      TabIndex        =   17
      Top             =   480
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "List files of type:"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   3180
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Drives:"
      Height          =   195
      Left            =   2460
      TabIndex        =   15
      Top             =   2295
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Foldername:"
      Height          =   195
      Left            =   2460
      TabIndex        =   14
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folders:"
      Height          =   195
      Left            =   2460
      TabIndex        =   13
      Top             =   135
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File name:"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddFlags As Long
Dim FilePath As String
Dim FileTypeFlag As Long


Private Sub chkRemoveLocal_Click()

    ' Set flags for removing local copy
    If chkRemoveLocal.Value = 1 Then
        AddFlags = AddFlags + VSSFLAG_DELYES
    Else
        AddFlags = AddFlags - VSSFLAG_DELYES
    End If
    
End Sub

Private Sub cmbFileType_Click()

    AddFlags = AddFlags - FileTypeFlag

    Select Case cmbFileType.Text
        Case "Autodetect"
            FileTypeFlag = VSSFLAG_BINTEST
        Case "Binary"
            FileTypeFlag = VSSFLAG_BINBINARY
        Case "Text"
            FileTypeFlag = VSSFLAG_BINTEXT
    End Select
    
    AddFlags = AddFlags + FileTypeFlag
    
End Sub

Private Sub cmdClose_Click()
    
    ' Close this form
    Unload Me
    
End Sub

Private Sub cmdAdd_Click()

    Dim Count As Integer
    Dim Response As Long
    Dim objVSSProject As VSSItem
    Dim Nodex As Node
    Dim AddingFromRoot As Boolean
    Dim objListItem As ListItem
    Dim objVSSVersion As VSSVersion
    Dim objVSSObject As VSSItem
    Dim FileDate As String
    Dim Selected As Boolean
    Dim ProjectName As String
    Dim ProjectToMoveTo As String
    
    ' Set Mouespointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    'Initialize variables
    Selected = False
    
    ' Set Add Flags
    If chkStoreLatestVersion.Value = 1 Then
        AddFlags = AddFlags + VSSFLAG_DELTANO
    Else
        AddFlags = AddFlags + VSSFLAG_DELTAYES
    End If
    
    ' Check if adding files to root project and set VSSItem
    ' as appropriate
    If frmMain.TreeView1.SelectedItem.Key = "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
    End If
    
    ' Adding files
    If chkRecursive.Enabled = False Then
    
        ' Adding files from the root of a drive
        If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then AddingFromRoot = True
        
        ' Iterate through the File List
        For Count = 1 To ListView1.ListItems.Count
                
            ' The file is selected
            If ListView1.ListItems.Item(Count).Selected Then
                    
                ' A file is selected so set flag
                Selected = True
                    
                ' Set variable FilePath to path of file to add
                If AddingFromRoot Then
                    FilePath = Dir1.Path + ListView1.ListItems.Item(Count).Key
                Else
                    FilePath = Dir1.Path + "\" + ListView1.ListItems.Item(Count).Key
                End If
                    
                Call AddFileToSourceSafe(FilePath, txtComment, AddFlags)
                ListView1.ListItems.Remove (ListView1.ListItems(Count))
                
            End If
        Next Count
        
        ' If selected is false (no files selected in filelist) then check to see
        ' if user typed in a file name
        If Not Selected And txtFileName <> "" Then
            
            ' Set the filepath and add file to VSS
            FilePath = txtFileName.Text
            Call AddFileToSourceSafe(FilePath, txtComment, AddFlags)
        End If
    
    ' Adding a folder
    Else
               
        ' Set flags and Add Folder
        If chkRecursive.Value = 1 Then
            AddFlags = AddFlags + VSSFLAG_RECURSYES
        Else
            AddFlags = AddFlags + VSSFLAG_RECURSNO
        End If
        ProjectName = GetFileName(txtFileName)
        objVSSProject.NewSubproject Name:=ProjectName, Comment:=txtComment.Text
        If Right(objVSSProject.Spec, 1) = "/" Then
            ProjectName = objVSSProject.Spec + ProjectName
        Else
            ProjectName = objVSSProject.Spec + "/" + ProjectName
        End If
        Set objVSSProject = Nothing
        Set objVSSProject = objVSSDatabase.VSSItem(ProjectName, False)
        objVSSProject.Add Local:=Dir1.Path, Comment:=txtComment.Text, iFlags:=AddFlags
        
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
        If InStr(3, ProjectName, "/") <> 0 Then
            ProjectToMoveTo = Left(ProjectName, InStr(3, ProjectName, "/") - 1)
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
        End If
        While InStr(Len(ProjectToMoveTo) + 2, ProjectName, "/") <> 0
            ProjectToMoveTo = Left(ProjectName, InStr(Len(ProjectToMoveTo) + 2, ProjectName, "/") - 1)
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
            frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
        Wend
        
        ProjectToMoveTo = ProjectName
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
        
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:

        ' If error is from a non unique key we can ignore it or, since
        ' we are removing an item from the ListView control we may get
        ' an index out of bounds error
        If Err <> 35602 And Err <> 35600 Then
            Response = MsgBox("Unable to add item." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
End Sub

Private Sub Dir1_Change()
    
    lstFiles.Path = Dir1.Path
    txtFileName.Text = Dir1.Path

End Sub

Private Sub Dir1_GotFocus()
    
    chkRecursive.Enabled = True
    txtFileName.Text = Dir1.Path
    cmdAdd.Enabled = True
    
End Sub

Private Sub Drive1_Change()

    Dim Response As Long
    
    ' Set On Error Rouine
    On Error GoTo ErrHandler
    
    Dir1.Path = Drive1.Drive
        
    If Err <> 0 Then
    
ErrHandler:

        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
    
    End If
End Sub

Private Sub ListView1_Click()
    
    Dim Count As Integer
    txtFileName.Text = ""
    
    ' Populate the File Text box with selected files
    For Count = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(Count).Selected Then
            txtFileName.Text = txtFileName.Text + ListView1.ListItems.Item(Count).Text + " "
        End If
    Next
    chkRecursive.Enabled = False
    
End Sub

Private Sub Form_Load()

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize form
    Flags = 0
    lstFiles.Pattern = "*.*"
    cmbFileType.AddItem ("Autodetect")
    cmbFileType.AddItem ("Binary")
    cmbFileType.AddItem ("Text")
    cmbFileType.ListIndex = 0
    FileTypeFlag = VSSFLAG_BINTEST
    
    Dim objFileListItem As ListItem
    Dim Count As Integer
    
    For Count = 0 To lstFiles.ListCount - 1
        Set objFileListItem = ListView1.ListItems.Add(, lstFiles.List(Count), lstFiles.List(Count))
    Next
    
End Sub

Private Sub lstFiles_PathChange()

    Dim Count As Integer
    Dim objFileListItem As ListItem
        
    ListView1.ListItems.Clear
    For Count = 0 To lstFiles.ListCount - 1
        Set objFileListItem = ListView1.ListItems.Add(, lstFiles.List(Count), lstFiles.List(Count))
    Next
    

End Sub

Private Sub txtFileName_Change()
    
    ' Enable ADD command button if there are files to add
    If txtFileName <> "" Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub txtFilePattern_Change()
    lstFiles.Pattern = txtFilePattern.Text
End Sub

Public Sub AddFileToSourceSafe(FilePath As String, Comment As String, AddFlags As Long)
    
    Dim FileName As String
    Dim objListItem As ListItem
    Dim Response As Long
    
    ' Set on error routine
    On Error GoTo ErrHandler
    
    ' Get the Filename from FilePath
    FileName = FilePath
    FileName = GetFileName(FileName)
    
    ' Add the file to the VSS database
    objVSSProject.Add Local:=FilePath, Comment:=Comment, iFlags:=AddFlags

    ' Add file to the File List
    Set objListItem = frmMain.ListView1.ListItems.Add(, frmMain.TreeView1.SelectedItem.Key + "/" + FileName, FileName, , "Leaf")
    
    ' Get the file date
    Set objVSSObject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key + "/" + FileName, False)
    For Each objVSSVersion In objVSSObject.Versions
        FileDate = objVSSVersion.Date
        Exit For
    Next
    objListItem.SubItems(2) = FileDate
    frmMain.ListView1.Sorted = True
    frmMain.ListView1.Refresh
        
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' If error is from a non unique key we can ignore it
        If Err <> 35602 Then
            Response = MsgBox("Unable to Add File." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    End If
End Sub
