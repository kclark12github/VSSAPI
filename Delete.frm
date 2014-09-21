VERSION 5.00
Begin VB.Form frmDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete"
   ClientHeight    =   1365
   ClientLeft      =   3330
   ClientTop       =   4245
   ClientWidth     =   5055
   Icon            =   "Delete.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1365
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3900
      TabIndex        =   3
      Top             =   540
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3900
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
   Begin VB.CheckBox chkDestroy 
      Caption         =   "&Destroy permanently"
      Height          =   270
      Left            =   660
      TabIndex        =   1
      Top             =   795
      Width           =   2415
   End
   Begin VB.Label lblItem 
      Caption         =   "Item:"
      Height          =   540
      Left            =   270
      TabIndex        =   0
      Top             =   165
      Width           =   3540
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Destroy As Boolean

' Set Destory variable as appropriate

Private Sub chkDestroy_Click()
    If chkDestroy.Value = 1 Then
        Destroy = True
    Else
        Destroy = False
    End If
End Sub

Private Sub cmdOK_Click()
    
    Dim Response As Long
    Dim RetVal As Long
    Dim Count As Integer
    Dim objVSSFile As VSSItem

    ' Set on error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    Response = vbYes
    
    ' Hide the Delete dialog
    frmDelete.Hide
    
    ' Deleting a project
    If Selected = "Treeview1" Then

        If Destroy Then
                
            If WarnDestroy Then
                RetVal = MsgBox("Destroy cannot be undone; information will be lost permanently!" + vbCrLf + "Continue anyway?", vbQuestion + vbYesNo, AppTitle)
            Else
                RetVal = vbYes
            End If
            If RetVal = vbYes Then
            
                ' Delete and destroy Project
                objVSSProject.Destroy
            
                ' Update main window
                frmMain.TreeView1.Nodes.Remove (frmMain.TreeView1.SelectedItem.Key)
                frmMain.ListView1.ListItems.Clear
            End If

        ' Delete Project
        Else
            
            If WarnDelete Then
                RetVal = MsgBox("Delete all specified items?", vbQuestion + vbYesNo, AppTitle)
            Else
                RetVal = vbYes
            End If
            If RetVal = vbYes Then
                
                ' Delete Project but do not destroy
                objVSSProject.Deleted = True
                
                ' Update main window
                frmMain.TreeView1.Nodes.Remove (frmMain.TreeView1.SelectedItem.Key)
                frmMain.ListView1.ListItems.Clear
                frmMain.TreeView1.SelectedItem.Image = "Open"
                
                ' Populate the Project and File List
                Set objVSSProject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
                Call frmMain.PopulateMain(objVSSProject)
            End If
        End If
    
    ' Deleting a file
    Else
                
        ' Iterate through each item in the ListView Control
        For Count = 1 To frmMain.ListView1.ListItems.Count
            
            ' Item is selected
            If frmMain.ListView1.ListItems(Count).Selected = True Then
        
                ' Instantiate the item
                Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.ListItems(Count).Key, False)
        
                ' Check to see if file is checked out
                If objVSSFile.IsCheckedOut <> VSSFILE_NOTCHECKEDOUT Then Response = MsgBox("File '" + objVSSFile.Name + "' is checked out. Continue?", vbQuestion + vbYesNo, AppTitle)
                If Response = vbYes Then
        
                    ' Destroy file
                    If Destroy Then
                    
                        If WarnDestroy Then
                            RetVal = MsgBox("Destroy cannot be undone; information will be lost permanently!" + vbCrLf + "Delete " + objVSSFile.Name + " anyway?", vbQuestion + vbYesNo, AppTitle)
                        Else
                            RetVal = vbYes
                        End If
                        If RetVal = vbYes Then
        
                            ' Delete and destroy File
                            objVSSFile.Destroy
                        End If
            
                    ' Delete File
                    Else
                        
                        If WarnDelete Then
                            RetVal = MsgBox("Delete " + objVSSFile.Name + "?", vbQuestion + vbYesNo, AppTitle)
                        Else
                            RetVal = vbYes
                        End If
                        If RetVal = vbYes Then
                        
                            ' Delete File but do not destroy
                            objVSSFile.Deleted = True
                        End If
            
                    End If
                End If
                
                ' Reset Response in case user selected No
                Response = vbYes

            End If
        Next
    End If
        
    ' Close Form
    Unload Me
    
    ' Check for errors
    If Err <> 0 Then

ErrHandler:

        If Err <> 35602 Then
            Response = MsgBox("Unable to delete item!" + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    End If

End Sub

Private Sub cmdCancel_Click()

    ' Close Form
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Intialize variables
    Destroy = False
End Sub
