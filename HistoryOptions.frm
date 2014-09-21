VERSION 5.00
Begin VB.Form frmHistoryOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Options"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   Icon            =   "HistoryOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1751
      TabIndex        =   3
      Top             =   1290
      Width           =   1200
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "&Recursive"
      Height          =   315
      Left            =   232
      TabIndex        =   1
      Top             =   600
      Width           =   2760
   End
   Begin VB.CheckBox cmdIncludeHistory 
      Caption         =   "&Include file histories"
      Height          =   315
      Left            =   232
      TabIndex        =   0
      Top             =   270
      Width           =   2760
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   274
      TabIndex        =   2
      Top             =   1290
      Width           =   1200
   End
End
Attribute VB_Name = "frmHistoryOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdOK_Click()

    Dim VersionCount As Integer
    Dim Response As Long
    Dim ItemName As String
    Dim Count As Integer
    Dim objHistoryItem As ListItem
    Dim objVersionsCollection As IVSSVersions
    Dim HistoryFlags As Long
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Initialize variables
    VersionCount = 0
    HistoryFlags = 0
    
    ' Set Error Handler
    On Error GoTo ErrHandler
    
    ' Initialize History dialog
    frmHistory.ListView1.SmallIcons = frmMain.ImageList1
    frmHistoryOptions.Hide
    
    ' Set the flags
    If chkRecursive.Value = 1 Then
        HistoryFlags = VSSFLAG_RECURSYES
    Else
        HistoryFlags = VSSFLAG_RECURSNO
    End If
    If cmdIncludeHistory.Value = 0 Then HistoryFlags = HistoryFlags + VSSFLAG_HISTIGNOREFILES
    
    ' Show Project History
    If Selected = "Treeview1" Then
    
        frmHistory.cmdGet.Visible = False
        
        ' Instantiate the current VSS Item as appropriate
        If frmMain.TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSObject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key, False)
            frmHistory.Caption = "History of " + frmMain.TreeView1.SelectedItem.Key
        Else
            Set objVSSObject = objVSSDatabase.VSSItem(frmMain.TreeView1.SelectedItem.Key + "/", False)
            frmHistory.Caption = "History of " + frmMain.TreeView1.SelectedItem.Key + "/"
        End If
        
        ' Initialize History Dialog box
        frmHistory.ListView1.ColumnHeaders(1).Text = "Name"
        frmHistory.ListView1.ColumnHeaders(1).Width = 1000
        
        ' Instantiate the versions collection
        Set objVersionsCollection = objVSSObject.Versions(HistoryFlags)
        
        ' Iterate through the versions
        For Each objVSSVersion In objVersionsCollection
        
            ' Current version is a Label
            If Left(objVSSVersion.Action, 5) = "Label" Then
                
                ' Add item to History dialog
                Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, , objVSSVersion.VSSItem.Name, , "Labeled")
            
            ' Current version is a Share
            ElseIf Left(objVSSVersion.Action, 5) = "Share" Then
                For Count = 1 To Len(objVSSVersion.Action)
                    If Left(ItemName, 1) = "/" Then
                        ItemName = Right(objVSSVersion.Action, Count - 2)
                        Exit For
                    End If
                    ItemName = Right(objVSSVersion.Action, Count)
                Next
                
                ' Add item to History dialog
                Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, , ItemName)
            
            ' Current Version is not a Share or Label
            Else
                
                ' Add item to History dialog
                If objVSSObject.Spec = "$/" Then
                    Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, , "$/")
                Else
                    Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, , objVSSVersion.VSSItem.Name)
                End If
            End If
            
            ' Add data on current version to History dialog (User, Date and Action)
            objHistoryItem.SubItems(1) = objVSSVersion.UserName
            objHistoryItem.SubItems(2) = Str(objVSSVersion.Date)
            objHistoryItem.SubItems(3) = objVSSVersion.Action
            
            ' Tally total versions for project
            VersionCount = VersionCount + 1
        Next
        frmHistory.lblHistory.Caption = "History: " + Str(VersionCount) + " items"
        frmHistory.cmdDetails.Visible = False
        frmHistory.cmdView.Visible = False
        frmHistory.cmdDiff.Visible = False
        
        ' Show History dialog
        frmMain.StatusBar1.Panels(1).Text = "Ready"
        frmHistory.Show 1
    
    ' Show File History
    ElseIf Selected = "Listview1" Then
    
        ' Instantiate the selected file
        Set objVSSObject = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
        
        ' Set the dialog caption
        frmHistory.Caption = "History of " + objVSSObject.Name
        
        ' Iterate through each version of the selected file
        For Each objVSSVersion In objVSSObject.Versions
            
            ' If version is a Label then add label icon
            If Left(objVSSVersion.Action, 5) = "Label" Then
                
                ' Add item to History dialog
                Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, frmMain.ListView1.SelectedItem.Key + " " + Trim(Str(objVSSVersion.VersionNumber)), "", , "Labeled")
            Else
                ' Add item to History dialog
                Set objHistoryItem = frmHistory.ListView1.ListItems.Add(, frmMain.ListView1.SelectedItem.Key + " " + Trim(Str(objVSSVersion.VersionNumber)), Str(objVSSVersion.VersionNumber))
            End If
            
            ' Add data on current version to History dialog (User, Date and Action)
            objHistoryItem.SubItems(1) = objVSSVersion.UserName
            objHistoryItem.SubItems(2) = Str(objVSSVersion.Date)
            objHistoryItem.SubItems(3) = objVSSVersion.Action
            
            ' Tally total versions for selected file
            VersionCount = VersionCount + 1
        Next
        
        ' Set dialog caption
        frmHistory.lblHistory.Caption = "History: " + Str(VersionCount) + " items"
        
        ' Show History dialog
        frmMain.StatusBar1.Panels(1).Text = "Ready"
        frmHistory.Show 1
    
    ' No pane is selected
    Else
        Response = MsgBox("No file or project selected", vbExclamation, AppTitle)
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:

        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
            Err.Clear
        Else
            Err.Clear
            Resume Next
        End If
    End If
    
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
    
    ' Set Mousepointer
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    MousePointer = vbNormal
    
End Sub
