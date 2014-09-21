VERSION 5.00
Begin VB.Form frmCheckIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check In"
   ClientHeight    =   2985
   ClientLeft      =   3285
   ClientTop       =   5985
   ClientWidth     =   6525
   ClipControls    =   0   'False
   Icon            =   "Checkin.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive"
      Height          =   270
      Left            =   105
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.CommandButton cmdDiff 
      Caption         =   "&Diff"
      Height          =   330
      Left            =   5400
      TabIndex        =   7
      Top             =   1050
      Width           =   1035
   End
   Begin VB.TextBox txtCheckInFrom 
      Height          =   285
      Left            =   615
      TabIndex        =   0
      Top             =   165
      Width           =   4560
   End
   Begin VB.CheckBox chkRemoveLocalCopy 
      Caption         =   "Remove &local copy"
      Height          =   270
      Left            =   105
      TabIndex        =   2
      Top             =   837
      Width           =   2880
   End
   Begin VB.CheckBox chkKeepCheckedOut 
      Caption         =   "&Keep checked out"
      Height          =   270
      Left            =   105
      TabIndex        =   1
      Top             =   550
      Width           =   2880
   End
   Begin VB.TextBox txtComment 
      Height          =   1170
      Left            =   105
      TabIndex        =   4
      Top             =   1710
      Width           =   6330
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   5400
      TabIndex        =   5
      Top             =   165
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   5400
      TabIndex        =   6
      Top             =   615
      Width           =   1035
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      Caption         =   "From:"
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   165
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   1440
      Width           =   705
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objVSSFile As VSSItem

Private Sub chkRemoveLocalCopy_Click()
    
    ' If the user wants to remove local copy then don't
    ' let them keep it checked out
    If chkRemoveLocalCopy.Value = 1 Then
        chkKeepCheckedOut.Value = 0
        chkKeepCheckedOut.Enabled = False
    Else
        chkKeepCheckedOut.Enabled = True
    End If
    
End Sub

Private Sub cmdDiff_Click()
    
    Dim FilesDiffer As Boolean
    Dim FilePath As String
    Dim Response As Long
    Dim FileType As String
    
    ' Set On Error routine
    On Error GoTo ErrHandler

    ' Check to see if local copy has changed
    FilePath = txtCheckInFrom.Text
    If Trim(FilePath) = "" Then
        Response = MsgBox("Please enter a valid path to the local file.", vbExclamation, AppTitle)
        txtCheckInFrom.SetFocus
    Else
        If Right(txtCheckInFrom, 1) <> "\" Then FilePath = FilePath + "\"
        Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
        If objVSSFile.Binary Then
            FileType = "binary"
        Else
            FileType = "text"
        End If
        FilesDiffer = objVSSFile.IsDifferent(Local:=FilePath + objVSSFile.Name)
        If FilesDiffer Then
            Response = MsgBox("Local " + FileType + " file '" + objVSSFile.Name + "' is different than SourceSafe version.", vbInformation, AppTitle)
        Else
            Response = MsgBox("Local " + FileType + " file '" + objVSSFile.Name + "' is identical to SourceSafe version.", vbInformation, AppTitle)
        End If
    End If
    
    ' Check for Errors
    If Err <> 0 Then
ErrHandler:

        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
    End If
    
End Sub

Private Sub cmdOK_Click()

    Dim Count As Integer
    Dim Response As Long
    Dim CheckOutFolder As String
    Dim CheckinFlags As Long
    Dim lpBuffer As String * 100
    Dim TempPath As String
    Dim TempPathLength As Integer
        
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    lpBuffer = String(100, Chr(0))

    ' Get path to Windows Temp Directory
    TempPathLength = GetTempPath(100, lpBuffer)
    TempPath = Mid(lpBuffer, 1, TempPathLength)
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath + "\"

    ' Set Mouespointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Hide the Form
    frmCheckIn.Hide
    
    ' Set Flags for CheckIn
    If chkKeepCheckedOut.Value = 1 Then
        CheckinFlags = VSSFLAG_KEEPYES
    End If
    If chkRemoveLocalCopy.Value = 1 Then
        CheckinFlags = CheckinFlags + VSSFLAG_DELYES
    End If
    If chkRecursive.Value = 1 Then
        CheckinFlags = CheckinFlags + VSSFLAG_RECURSYES
    End If
    
    ' Checking In a file(s)
    If Selected = "Listview1" Then
    
        ' Iterate through File List
        For Count = 1 To frmMain.ListView1.ListItems.Count
        
            ' File is selected
            If frmMain.ListView1.ListItems(Count).Selected = True Then
            
                ' Instantiate VSSItem
                Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(Count).Key, False)
                
                ' Check the CheckOut status of the file
                Select Case objVSSFile.IsCheckedOut
                    
                    ' File is checked out by user
                    Case VSSFILE_CHECKEDOUT_ME
                    
                        ' Check if user has specified a CheckIn location
                        ' other than the CheckOut Folder
                        If Trim(txtCheckInFrom.Text) <> "" Then
                            CheckOutFolder = txtCheckInFrom.Text
                        Else
                        
                            ' Find where it is checked out to
                            For Each objVSSCheckout In objVSSFile.Checkouts
                                CheckOutFolder = objVSSCheckout.LocalSpec
                                Exit For
                            Next
                        End If
                        
                        ' CheckIn the file
                        If Right(CheckOutFolder, 1) <> "\" Then CheckOutFolder = CheckOutFolder + "\"
                        
                        Call CheckinFile(CheckOutFolder, TempPath, CheckinFlags + ForceDirFlag + CheckInUnchangedFlag)
                        
                    ' File is checked out by another user
                    Case VSSFILE_CHECKEDOUT
                        
                        ' If user is Admin then allow CheckIn of file
                        If UCase(objVSSDatabase.UserName) <> "ADMIN" Then
                            
                            Response = MsgBox("The file " + objVSSFile.Name + " is checked out to another user.", vbExclamation, AppTitle)
                        
                        Else
                        
                            For Each objVSSCheckout In objVSSFile.Checkouts
                                
                                ' Verify each CheckOut
                                Response = MsgBox("The file " + objVSSFile.Name + " is checked out by user '" + objVSSCheckout.UserName + "' to " + objVSSCheckout.LocalSpec + "'." + vbCrLf + "As user '" + UserName + "' you are allowed to complete this request. Continue?", vbQuestion + vbYesNo, AppTitle)
                                If Response = vbYes Then
                                    
                                    ' Check In the file
                                    CheckOutFolder = objVSSCheckout.LocalSpec
                                    If Right(CheckOutFolder, 1) <> "\" Then CheckOutFolder = CheckOutFolder + "\"
                                    If Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") <> "" Then
                                        Response = MsgBox("Have you resolved all conflicts?", vbQuestion + vbYesNo, AppTitle)
                                    Else
                                        Response = vbYes
                                    End If
                                    If Response = vbYes Then
                                        If Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") <> "" Then Kill CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org"
                                        objVSSFile.CheckIn Comment:=txtComment.Text, Local:=CheckOutFolder + objVSSFile.Name, iFlags:=CheckinFlags + ForceDirFlag + CheckInUnchangedFlag
                                    End If
                                    
                                    ' Check to see if a merge occured
                                    If Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") <> "" And Response <> vbNo Then
                                        Response = MsgBox("The file was merged and conflicts exist. Edit " + CheckOutFolder + objVSSFile.Name + " to resolve these before checking in.", vbExclamation, AppTitle)
                                    End If
                                End If
                            Next
                        End If
                
                    ' File is not checked out
                    Case VSSFILE_NOTCHECKEDOUT
                
                        Response = MsgBox("You do not have the file " + objVSSFile.Name + " checked out.", vbExclamation, AppTitle)
                
                End Select
            End If
            
ContinueAfterError:

        Next
        
        ' Checking in a Project
        Else
            
            ' Instantiate VSSItem
            If frmMain.TreeView1.SelectedItem.Key = "$" Then
                Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
            Else
                Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
            End If

            ' Check to see where to check in from
            If txtCheckInFrom.Text <> "" Then
               
               ' CheckIn the Project
                objVSSProject.CheckIn Comment:=txtComment.Text, Local:=txtCheckInFrom.Text, iFlags:=CheckinFlags + ForceDirFlag + CheckInUnchangedFlag
            Else
                
                ' CheckIn the Project
                objVSSProject.CheckIn Comment:=txtComment.Text, iFlags:=CheckinFlags + ForceDirFlag + CheckInUnchangedFlag
            End If
            frmMain.txtResults.Text = frmMain.txtResults.Text + "Project " + objVSSProject.Name + " checked in" + vbCrLf
            Call ShowResults

        End If
    
     ' Check for errors
     If Err <> 0 Then
ErrHandler:
        
        Response = MsgBox("Unable to CheckIn file '" + objVSSFile.Name + "'." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
        If Selected = "Listview1" Then Resume ContinueAfterError
    End If
    
    ' Update GUI
    frmMain.RefreshFileList
    
    ' Set Mouespointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    
    ' Close Form
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' Center form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

End Sub

' Checks in the file and opens the Visual Merge dialog as required
Public Sub CheckinFile(CheckOutFolder As String, TempPath As String, Flags As Long)

    Dim objCheckOut As VSSCheckout
    Dim objBaseFile As VSSItem
    Dim Response As Long
    Dim MyFile As String
    Dim BaseFile As String
    Dim LatestFile As String
    Dim NewFile As String
    
    ' Set the local, latest and base files
    MyFile = objVSSFile.LocalSpec
    LatestFile = TempPath + objVSSFile.Name
    objVSSFile.Get Local:=LatestFile, iFlags:=VSSFLAG_REPREPLACE
    NewFile = TempPath + objVSSFile.Name + ".New"
    If Dir(NewFile) <> "" Then Kill NewFile
    For Each objCheckOut In objVSSFile.Checkouts
        If objCheckOut.UserName = UserName And objCheckOut.LocalSpec = GetDirPath(objVSSFile.LocalSpec) Then
            Set objBaseFile = objVSSFile.Version(objCheckOut.VersionNumber)
            BaseFile = TempPath + "BaseFile"
            objBaseFile.Get Local:=BaseFile, iFlags:=VSSFLAG_REPREPLACE
            Exit For
        End If
    Next
    
    ' Open Visual Merge dialog if needed. If <filename>.org exists we
    ' know that the merge took place and the file was saved but the checkin
    ' was postponed so check in now and skip the Visual Merge.
    If Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") = "" And DiffMethod = True Then
        frmMain.VisMerge1.DoVisMerge MyFile, BaseFile, LatestFile, NewFile, frmMain.hWnd
    End If
    MyFile = Dir(NewFile)
    
    ' If NewFile exists then a merge took place
    If MyFile = "" Then
    
        ' Check to see if we are checking in an uncompleted merge
        If Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") <> "" Then
            Response = MsgBox("Have you resolved all conflicts?", vbQuestion + vbYesNo, AppTitle)
        Else
            Response = vbYes
        End If
    
        ' Check In then file
        If Response = vbYes Then
            objVSSFile.CheckIn Comment:=txtComment.Text, Local:=CheckOutFolder + objVSSFile.Name, iFlags:=Flags
            
            ' Check to see if a merge occured
            If Not DiffMethod And Dir(CheckOutFolder + GetFileNameNoExt(objVSSFile.Name) + ".org") <> "" Then
                Response = MsgBox("The file was merged and conflicts exist. Edit " + CheckOutFolder + objVSSFile.Name + " to resolve these before checking in.", vbExclamation, AppTitle)
            Else
                frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " checked in" + vbCrLf
                Call ShowResults
            End If
        End If
    Else
    
        ' Call the checkin method to create a local version
        ' of the merge file. This does not check the file in
        objVSSFile.CheckIn Comment:=txtComment.Text, Local:=CheckOutFolder + objVSSFile.Name, iFlags:=Flags

        ' If user said "no" to saving the file then the filelength of NewFile is 0
        If FileLen(NewFile) = 0 Then
        
            Response = MsgBox("Edit the file '" + CheckOutFolder + objVSSFile.Name + "' to resolve any conflicts before checking in.", vbExclamation, AppTitle)
            
        ' User saved the file so it should already be checked in
        ElseIf GetAttr(CheckOutFolder + objVSSFile.Name) <> vbReadOnly + vbArchive Then
            
            ' Copy the user's merges over the local copy
            If Dir(NewFile) <> "" Then
                FileCopy NewFile, CheckOutFolder + objVSSFile.Name
            End If
        
            ' Ask if user wants to check in merged file
            Response = MsgBox("Do you want to check the file in now?", vbQuestion + vbYesNo, AppTitle)
            If Response = vbYes Then
                objVSSFile.CheckIn Comment:=txtComment.Text, Local:=CheckOutFolder + objVSSFile.Name, iFlags:=Flags
                frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " checked in" + vbCrLf
                Call ShowResults
            End If
        End If
        
        ' Delete the temp copy of the merged file
        If Dir(NewFile) <> "" Then Kill NewFile

    End If
    
End Sub
