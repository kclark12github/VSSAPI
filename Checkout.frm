VERSION 5.00
Begin VB.Form frmCheckOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Out"
   ClientHeight    =   2685
   ClientLeft      =   3825
   ClientTop       =   3375
   ClientWidth     =   6525
   ClipControls    =   0   'False
   Icon            =   "Checkout.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkExclusive 
      Caption         =   "&Allow multiple checkouts"
      Height          =   330
      Left            =   1020
      TabIndex        =   3
      Top             =   1793
      Value           =   1  'Checked
      Width           =   2310
   End
   Begin VB.ComboBox cmbReplaceWriteable 
      Height          =   315
      Left            =   3525
      TabIndex        =   6
      Top             =   2250
      Width           =   1890
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive"
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ComboBox cmbFileTime 
      Height          =   315
      Left            =   3525
      TabIndex        =   5
      Top             =   1620
      Width           =   1890
   End
   Begin VB.CheckBox chkDontGetLocal 
      Caption         =   "&Don't Get local copy"
      Height          =   330
      Left            =   1020
      TabIndex        =   2
      Top             =   1470
      Width           =   2310
   End
   Begin VB.TextBox txtCheckOutTarget 
      Height          =   330
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   4275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Top             =   525
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txtComment 
      Height          =   795
      Left            =   1020
      TabIndex        =   1
      Top             =   525
      Width           =   4275
   End
   Begin VB.Label lblReplaceWriteable 
      AutoSize        =   -1  'True
      Caption         =   "Replace writable:"
      Height          =   195
      Left            =   3525
      TabIndex        =   12
      Top             =   2010
      Width           =   1230
   End
   Begin VB.Label lblFileTime 
      AutoSize        =   -1  'True
      Caption         =   "Set file time"
      Height          =   195
      Left            =   3525
      TabIndex        =   11
      Top             =   1380
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   105
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   570
      Width           =   705
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WorkingFolder As String
Dim FileTimeFlag As Long
Dim ReplaceWriteableFlag As Long
Dim CheckOutStatus As Long

Private Sub chkExclusive_Click()

    ' Set the CheckOut Status flag
    If chkExclusive.Value = 0 Then
        CheckOutStatus = VSSFLAG_CHKEXCLUSIVEYES
    Else
        CheckOutStatus = VSSFLAG_CHKEXCLUSIVENO
    End If
    
End Sub

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
    Dim FileMayBeCheckedOut As Boolean
    Dim TempFlag As Long
    
    ' Set Mouespointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Initialize variables
    Flags = 0
    FileMayBeCheckedOut = False
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set Get Local Copy flag
    If chkDontGetLocal.Value = 1 Then Flags = VSSFLAG_GETNO
    
    ' Set Replace Writeable flag
    Flags = Flags + ReplaceWriteableFlag
    
    ' Set CheckOut Time flag
    Flags = Flags + FileTimeFlag
    
    ' Set CheckOut status flag
    Flags = Flags + CheckOutStatus
    
    ' Get current date and time
    FileDate = Date + Time
    
    ' Hide the Form
    frmCheckOut.Hide
    
    ' Checking out a file(s)
    If Selected = "Listview1" Then
    
        ' Iterate through all files in ListView
        For Count = 1 To frmMain.ListView1.ListItems.Count
            
            ' File is selected
            If frmMain.ListView1.ListItems(Count).Selected = True Then

                ' Instantiate VSSItem
                Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(Count).Key, False)
                
                ' Set Checkout target
                If frmCheckOut.txtCheckOutTarget.Text <> "" Then
                    If Right(frmCheckOut.txtCheckOutTarget.Text, 1) = "\" Then
                        WorkingFolder = frmCheckOut.txtCheckOutTarget.Text + objVSSFile.Name
                    Else
                        WorkingFolder = frmCheckOut.txtCheckOutTarget.Text + "\" + objVSSFile.Name
                    End If
                Else
                    If Right(WorkingFolder, 1) = "\" Then
                        WorkingFolder = WorkingFolder + objVSSFile.Name
                    Else
                        WorkingFolder = WorkingFolder + "\" + objVSSFile.Name
                    End If
                End If
                        
                ' Validate the CheckOut status of the file
                Select Case objVSSFile.IsCheckedOut
                    
                    ' File is checked out by user
                    Case VSSFILE_CHECKEDOUT_ME
                        Response = MsgBox("You currently have the file " + objVSSFile.Name + " checked out.", vbExclamation, AppTitle)
                        FileMayBeCheckedOut = False
                        
                    ' File is checked out by another user
                    Case VSSFILE_CHECKEDOUT
                    
                        If WarnCheckOut Then
                            Response = MsgBox("The file " + objVSSFile.Name + " is checked out to another user. Continue?", vbExclamation + vbYesNo, AppTitle)
                        Else
                            Response = vbYes
                        End If
                        FileMayBeCheckedOut = (Response = vbYes)
                        
                    ' File is not checked out
                    Case Else
                        
                        FileMayBeCheckedOut = True
                        
                 End Select
    
                ' Check the file out
                If FileMayBeCheckedOut Then
                
                    ' Check if user wants to be warned about writeable copies
                    TempFlag = 0
                    If ReplaceWriteableFlag = VSSFLAG_REPASK Then
                        If GetAttr(WorkingFolder) = vbArchive Then
                            If Err <> 53 Then
                                Response = MsgBox("File '" + WorkingFolder + "' is writeable. Replace?", vbQuestion + vbYesNo, AppTitle)
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
                    objVSSFile.CheckOut Comment:=frmCheckOut.txtComment.Text, Local:=WorkingFolder, iFlags:=Flags + ForceDirFlag + EOLFlag
                    frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " checked out" + vbCrLf
                    Call ShowResults
                End If
                
                ' Reset Replace Writeable Flag
                If TempFlag <> 0 Then
                    Flags = Flags - TempFlag
                    Flags = Flags + VSSFLAG_REPASK
                End If
            End If
        Next
    
    ' Checking out a project
    Else
        
        ' Instantiate VSSItem
        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
        End If
        
        ' Set recursive flags
        If frmCheckOut.chkRecursive.Value = 1 Then
            Flags = Flags + VSSFLAG_RECURSYES
        Else
            Flags = Flags + VSSFLAG_RECURSNO
        End If
        
        ' Set Checkout target
        If frmCheckOut.txtCheckOutTarget.Text <> "" Then
            
            ' CheckOut the Project
            objVSSProject.CheckOut Comment:=txtComment.Text, Local:=txtCheckOutTarget.Text, iFlags:=Flags + ForceDirFlag + EOLFlag
        Else
            
            ' CheckOut the Project
            objVSSProject.CheckOut Comment:=txtComment.Text, iFlags:=Flags + ForceDirFlag + EOLFlag
        End If
        frmMain.txtResults.Text = frmMain.txtResults.Text + "Project " + objVSSProject.Name + " checked out" + vbCrLf
        Call ShowResults
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
    
        ' Local file not found when checking for writeable copy
        If Err = 53 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
            
        ' File is checked out by another user and is binary
         ElseIf Err = -2147166248 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
            
         ElseIf Err <> 35602 Then
            Response = MsgBox("Unable to CheckOut file(s)." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    End If
    
    ' Update GUI
    frmMain.RefreshFileList
    
    ' Set Mouespointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
    ' Close form
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    
    ' Close form
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim Response As Long
    Dim WorkingFolder As String

    ' Set On Error routine
    On Error Resume Next

    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize variables
    CheckOutStatus = VSSFLAG_CHKEXCLUSIVENO
    
    ' Set CheckOut text box to Working Folder
    If Selected = "Listview1" Then
        WorkingFolder = objVSSProject.Parent.LocalSpec
    Else
        WorkingFolder = objVSSProject.LocalSpec
    End If
    
    ' Verify working folder exists
    If WorkingFolder = "" Then
        Err.Clear
        Response = MsgBox("You must set a working folder for this file before it can be edited." + vbCrLf + "Would you like to set one now?", vbExclamation + vbYesNo, AppTitle)
        If Response = vbYes Then Call frmMain.SetWorkingFolder
        Unload Me
        Exit Sub
    Else
        txtCheckOutTarget.Text = WorkingFolder
    End If
    
    ' Working Folder Set
    If Err = 0 Then
    
        ' Initialize ComboBoxes and set Flags to 'Current' and 'Ask'
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
    End If
    
End Sub

