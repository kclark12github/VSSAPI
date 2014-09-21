VERSION 5.00
Begin VB.Form frmUndoCheckOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Undo Check Out"
   ClientHeight    =   1485
   ClientLeft      =   2580
   ClientTop       =   3555
   ClientWidth     =   3705
   ClipControls    =   0   'False
   Icon            =   "Undocheckout.frx":0000
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1485
   ScaleWidth      =   3705
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive"
      Height          =   285
      Left            =   165
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2580
      TabIndex        =   2
      Top             =   585
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2580
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.ComboBox cmbLocalCopy 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   555
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local copy:"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmUndoCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim FileCount As Integer
    Dim Response As Long
    Dim UndoCheckOutFlags As Long
    Dim objVSSFile As VSSItem
    Dim CheckOutFolder As String
    Dim CheckOutPath As String
    Dim RetVal As Long
    
    ' Set Mouespointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Initialize variables
    UndoCheckOutFlags = 0
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set UndoCheckOut Flag
    Select Case cmbLocalCopy.Text
    
        Case "Replace"
        
            UndoCheckOutFlags = VSSFLAG_REPREPLACE
            
        Case "Delete"
        
            UndoCheckOutFlags = VSSFLAG_DELYES
            
        Case "Leave"
        
            UndoCheckOutFlags = VSSFLAG_DELNOREPLACE
            
    End Select
        
    ' Set recursive flag
    If chkRecursive.Value = 1 Then
        UndoCheckOutFlags = UndoCheckOutFlags + VSSFLAG_RECURSYES
    Else
        UndoCheckOutFlags = UndoCheckOutFlags + VSSFLAG_RECURSNO
    End If
    
    ' Hide the Form
    frmUndoCheckOut.Hide
    
    ' UncheckOut a file(s)
    If Selected = "Listview1" Then

        ' Iterate through file list
        For FileCount = 1 To frmMain.ListView1.ListItems.Count
        
            ' File is selected
            If frmMain.ListView1.ListItems(FileCount).Selected = True Then
            
                ' Instantiate the file object
                Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(FileCount).Key, False)
                
                ' Verify file's CheckOut status
                Select Case objVSSFile.IsCheckedOut
                
                    ' File is checked out by user
                    Case VSSFILE_CHECKEDOUT_ME
                    
                        ' Get the CheckOut Folder and the File Date
                        For Each objVSSCheckOut In objVSSFile.Checkouts
                            CheckOutPath = objVSSCheckOut.LocalSpec
                            If Right(CheckOutPath, 1) <> "\" Then CheckOutPath = CheckOutPath + "\"
                            If objVSSCheckOut.UserName = UserName And CheckOutPath + objVSSFile.Name = objVSSFile.LocalSpec Then
                                CheckOutFolder = objVSSCheckOut.LocalSpec
                                If Right(CheckOutFolder, 1) = "\" Then
                                    CheckOutPath = CheckOutFolder + objVSSFile.Name
                                Else
                                    CheckOutPath = CheckOutFolder + "\" + objVSSFile.Name
                                End If
                                FileDate = objVSSCheckOut.Date
                                Exit For
                            End If
                        Next
                        
                        ' UnCheckOut the File
                        If WarnUndoCheckOut Then
                            If objVSSFile.IsDifferent(Local:=CheckOutPath) And cmbLocalCopy.Text <> "Leave" And cmbLocalCopy.Text <> "Delete" Then
                                RetVal = MsgBox(CheckOutPath + " has changed. Undo check out and lose changes?", vbQuestion + vbYesNo, AppTitle)
                            Else
                                RetVal = vbYes
                            End If
                        Else
                            RetVal = vbYes
                        End If
                        If RetVal = vbYes Then
                            If Right(CheckOutFolder, 1) = "\" Then
                                objVSSFile.UndoCheckOut Local:=CheckOutFolder + objVSSFile.Name, iFlags:=UndoCheckOutFlags + ForceDirFlag
                            Else
                                objVSSFile.UndoCheckOut Local:=CheckOutFolder + "\" + objVSSFile.Name, iFlags:=UndoCheckOutFlags + ForceDirFlag
                            End If
                            frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " unchecked out" + vbCrLf
                            Call ShowResults
                        End If
                        
                    ' File is checked out to another user
                    Case VSSFILE_CHECKEDOUT
                    
                        ' If user is Admin then allow UnCheckOut of file
                        If UCase(objVSSDatabase.UserName) <> "ADMIN" Then
                            Response = MsgBox("The file " + objVSSFile.Name + " is checked out to another user.", vbExclamation, AppTitle)
                        Else
                        
                            Response = MsgBox("The file " + objVSSFile.Name + " is checked out by another user. As user '" + UserName + "' you are allowed to complete this request. Continue?", vbQuestion + vbYesNo, AppTitle)
                            If Response = vbYes Then
                        
                                ' Get the CheckOut Folder
                                CheckOutFolder = objVSSCheckOut.LocalSpec
                                If Right(CheckOutFolder, 1) = "\" Then
                                    CheckOutPath = CheckOutFolder + objVSSFile.Name
                                Else
                                    CheckOutPath = CheckOutFolder + "\" + objVSSFile.Name
                                End If
                            
                                ' UnCheckOut the File
                                If WarnUndoCheckOut Then
                                    If objVSSFile.IsDifferent(Local:=CheckOutPath) Then
                                        RetVal = MsgBox(CheckOutPath + " has changed. Undo check out and lose changes?", vbQuestion + vbYesNo, AppTitle)
                                    Else
                                        RetVal = vbYes
                                    End If
                                Else
                                    RetVal = vbYes
                                End If
                                If RetVal = vbYes Then
                                    If Right(CheckOutFolder, 1) = "\" Then
                                        objVSSFile.UndoCheckOut Local:=CheckOutFolder + objVSSFile.Name, iFlags:=UndoCheckOutFlags + ForceDirFlag
                                    Else
                                        objVSSFile.UndoCheckOut Local:=CheckOutFolder + "\" + objVSSFile.Name, iFlags:=UndoCheckOutFlags + ForceDirFlag
                                    End If
                                    frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " unchecked out" + vbCrLf
                                    Call ShowResults
                                End If
                            End If
                        End If
                        
                    ' File is not checked out
                    Case VSSFILE_NOTCHECKEDOUT
                    
                        Response = MsgBox("You do not have the file " + objVSSFile.Name + " checked out.", vbExclamation, AppTitle)
                
                End Select
            End If
    
ContinueAfterError:
        Next
        
    ' Uncheckout a project
    Else
    
        ' Instantiate VSSItem
        If frmMain.TreeView1.SelectedItem.Key = "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
        End If
        
        ' Undo the checkout(s)
        objVSSProject.UndoCheckOut iFlags:=UndoCheckOutFlags + ForceDirFlag
        frmMain.txtResults.Text = frmMain.txtResults.Text + "Project " + objVSSProject.Name + " unchecked out" + vbCrLf
        Call ShowResults
    End If
    
     ' Check for errors
     If Err <> 0 Then
ErrHandler:
        
        Response = MsgBox("Unable to UnCheckOut file '" + objVSSFile.Name + "'." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
        If Selected = "Listview1" Then Resume ContinueAfterError
    End If
    
    ' Set Mouespointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
    ' Refresh GUI and Close Form
    frmMain.RefreshFileList
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Populate ComboBox
    cmbLocalCopy.AddItem "Replace"
    cmbLocalCopy.AddItem "Delete"
    cmbLocalCopy.AddItem "Leave"
    cmbLocalCopy.ListIndex = 0
    
End Sub
