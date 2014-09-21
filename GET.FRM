VERSION 5.00
Begin VB.Form frmGet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get"
   ClientHeight    =   1650
   ClientLeft      =   4320
   ClientTop       =   5730
   ClientWidth     =   6525
   ClipControls    =   0   'False
   Icon            =   "get.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1650
   ScaleWidth      =   6525
   Begin VB.CheckBox chkMakeWriteable 
      Caption         =   "&Make writeable"
      Height          =   255
      Left            =   615
      TabIndex        =   3
      Top             =   1275
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ComboBox cmbReplaceWriteable 
      Height          =   315
      Left            =   615
      TabIndex        =   1
      Top             =   825
      Width           =   1890
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive"
      Height          =   285
      Left            =   3030
      TabIndex        =   4
      Top             =   1275
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbFileTime 
      Height          =   315
      Left            =   3030
      TabIndex        =   2
      Top             =   825
      Width           =   1890
   End
   Begin VB.TextBox txtGetTarget 
      Height          =   300
      Left            =   615
      TabIndex        =   0
      Top             =   165
      Width           =   4275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   5400
      TabIndex        =   5
      Top             =   165
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   5400
      TabIndex        =   6
      Top             =   645
      Width           =   1035
   End
   Begin VB.Label lblReplaceWriteable 
      AutoSize        =   -1  'True
      Caption         =   "Replace writable"
      Height          =   195
      Left            =   615
      TabIndex        =   9
      Top             =   615
      Width           =   1185
   End
   Begin VB.Label lblFileTime 
      AutoSize        =   -1  'True
      Caption         =   "Set File time"
      Height          =   195
      Left            =   3030
      TabIndex        =   8
      Top             =   615
      Width           =   855
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   225
      Width           =   240
   End
End
Attribute VB_Name = "frmGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WorkingFolder As String
Dim FileTimeFlag As Long
Dim ReplaceWriteableFlag As Long

Private Sub cmbFileTime_Click()
    
    ' Set the file time flag
    Select Case cmbFileTime.Text
        Case "CheckIn"
            FileTimeFlag = VSSFLAG_TIMEUPD
        Case "Current"
            FileTimeFlag = VSSFLAG_TIMENOW
        Case "Modification"
            FileTimeFlag = VSSFLAG_TIMEMOD
    End Select
End Sub

Private Sub cmbReplaceWriteable_Click()

    ' Set the replace writeable flag
    Select Case cmbReplaceWriteable.Text
        Case "Ask"
            ReplaceWriteableFlag = VSSFLAG_REPASK
        Case "Merge"
            ReplaceWriteableFlag = VSSFLAG_REPMERGE
        Case "Replace"
            ReplaceWriteableFlag = VSSFLAG_REPREPLACE
        Case "Skip"
            ReplaceWriteableFlag = VSSFLAG_REPSKIP
        End Select
    
End Sub

Private Sub cmdOK_Click()
    
    Dim Count As Integer
    Dim Response As Long
    Dim objVSSFile As VSSItem
    Dim GetTarget As String
    Dim Flags As Long
    Dim TempFlag As Long
    
    ' Initialize variables
    Response = vbYes
    Flags = VSSFLAG_REPREPLACE
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set the File Time Flag
    Flags = FileTimeFlag
    
    ' Set Replace writeable flag
    Flags = Flags + ReplaceWriteableFlag
    
    ' Check for writeable
    frmGet.Hide
    
    ' Getting a file(s)
    If Selected = "Listview1" Then
    
        ' Iterate through the filelist and Get the selected files
        For Count = 1 To frmMain.ListView1.ListItems.Count
            
            ' File is selected
            If frmMain.ListView1.ListItems(Count).Selected = True Then
                
                ' Instantiate the file to be gotten
                Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(Count).Key, False)
                If Trim(txtGetTarget.Text) <> "" Then
                    If Mid(txtGetTarget.Text, Len(txtGetTarget.Text), 1) = "\" Then
                        GetTarget = txtGetTarget.Text + objVSSFile.Name
                    Else
                        GetTarget = txtGetTarget.Text + "\" + objVSSFile.Name
                    End If
                End If
                
                ' Check if user wants to be warned about writeable copies
                TempFlag = 0
                If ReplaceWriteableFlag = VSSFLAG_REPASK Then
                    If GetAttr(GetTarget) = vbArchive Then
                        If Err <> 53 Then
                            Response = MsgBox("File '" + GetTarget + "' is writeable. Replace?", vbQuestion + vbYesNo, AppTitle)
                            Flags = Flags - VSSFLAG_REPASK
                            If Response = vbNo Then
                                TempFlag = VSSFLAG_REPSKIP
                            Else
                                TempFlag = VSSFLAG_REPREPLACE
                            End If
                            Flags = Flags + TempFlag
                        End If
                    End If
                End If
                
LocalFileNotFound:
                
                ' Get the file(s)
                objVSSFile.Get Local:=GetTarget, iFlags:=Flags + ForceDirFlag + EOLFlag
                frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " gotten" + vbCrLf
                Call ShowResults
                    
                ' Make local copy writeable
                If chkMakeWriteable.Value = 1 Then
                    SetAttr GetTarget, vbArchive
                End If
            End If
        Next
        
    ' Getting a project
    Else
    
        ' Instantiate VSSItem
        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
        End If
        
        ' Set recursive flag
        If chkRecursive.Value = 1 Then Flags = Flags + VSSFLAG_RECURSYES

        ' Get the project
        If txtGetTarget.Text <> "" Then
            objVSSProject.Get Local:=txtGetTarget.Text, iFlags:=Flags + ForceDirFlag + EOLFlag
        Else
            objVSSProject.Get iFlags:=Flags + ForceDirFlag + EOLFlag
        End If
        frmMain.txtResults.Text = frmMain.txtResults.Text + "Project " + objVSSProject.Name + " gotten" + vbCrLf
        Call ShowResults
    
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:

        If Err = 53 And Selected = "Listview1" Then
            
            ' Local file not found when checking for writeable copy
            Err.Clear
            Resume LocalFileNotFound
        Else
            Response = MsgBox("Unable to Get file(s)." + vbCrLf + Err.Description, vbExclamation, AppTitle)
            Err.Clear
        End If
    End If
    
    ' Close the form
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    MousePointer = vbNormal
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()

    ' Close form
    Err.Clear
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Set On Error routine
    On Error Resume Next

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Set get text box to Working Folder (if there is one)
    txtGetTarget.Text = objVSSProject.Parent.LocalSpec
    
    ' Initialize File Time ComboBox and set Flags to 'Current'
    cmbFileTime.AddItem ("CheckIn")
    cmbFileTime.AddItem ("Current")
    cmbFileTime.AddItem ("Modification")
    cmbFileTime.ListIndex = 1
    FileTimeFlag = VSSFLAG_TIMENOW
    cmbReplaceWriteable.AddItem ("Replace")
    cmbReplaceWriteable.AddItem ("Skip")
    cmbReplaceWriteable.AddItem ("Merge")
    cmbReplaceWriteable.ListIndex = 1
    ReplaceWriteableFlag = VSSFLAG_REPSKIP
    
    ' Since automation doesn't support "Ask" for replace writeable
    ' we "Ask" this in VB code for file item only
    If Selected = "Listview1" Then cmbReplaceWriteable.AddItem ("Ask")
    
End Sub



