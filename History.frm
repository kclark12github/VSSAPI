VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History of"
   ClientHeight    =   4635
   ClientLeft      =   4980
   ClientTop       =   5970
   ClientWidth     =   7875
   ClipControls    =   0   'False
   Icon            =   "History.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   390
      Left            =   6810
      TabIndex        =   5
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdDiff 
      Caption         =   "Di&ff"
      Height          =   390
      Left            =   6810
      TabIndex        =   4
      Top             =   1845
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   390
      Left            =   6810
      TabIndex        =   3
      Top             =   1350
      Width           =   975
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details"
      Height          =   390
      Left            =   6810
      TabIndex        =   2
      Top             =   855
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   390
      Left            =   6810
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "VERSION"
         Object.Tag             =   ""
         Text            =   "Version"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "USER"
         Object.Tag             =   ""
         Text            =   "User"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "DATE"
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "ACTION"
         Object.Tag             =   ""
         Text            =   "Action"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Label lblHistory 
      AutoSize        =   -1  'True
      Caption         =   "History:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   525
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VersionNumberOne As Integer
Dim VersionNumberTwo As Integer
Dim MultipleFilesSelected As Boolean

Private Sub cmdClose_Click()

    ' Close Form
    Unload Me
    
End Sub

' Get the details on the selected version

Private Sub cmdDetails_Click()

    frmDetails.lblFile.Caption = "File: " + objVSSObject.Spec
    For Each objVSSVersion In objVSSObject.Versions
        If objVSSVersion.VersionNumber = Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " "))) Then
            frmDetails.lblDate.Caption = "Date: " + Str(objVSSVersion.Date)
            frmDetails.lblVersion.Caption = "Version: " + Str(objVSSVersion.VersionNumber)
            frmDetails.lblUser.Caption = "User: " + objVSSVersion.UserName
            frmDetails.txtLabel.Text = objVSSVersion.Label
            frmDetails.txtComment.Text = objVSSVersion.Comment
            frmDetails.txtAction.Text = objVSSVersion.Action
            If Left(objVSSVersion.Action, 5) = "Label" Then
                frmDetails.txtLabelComment.Text = objVSSVersion.LabelComment
            Else
                frmDetails.txtLabelComment.Enabled = False
                frmDetails.txtLabelComment.BackColor = &H8000000F
            End If
            Exit For
        End If
    Next
    
    ' Show the Detail Information
    frmDetails.Show 1
    
End Sub

Private Sub cmdGet_Click()
    
    Dim objVSSVersion As VSSItem
    Dim Response As Long
    Dim GetPath As String
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Check for a working folder
    If objVSSObject.LocalSpec = "" Then
        Response = MsgBox("Sorry, this command cannot be completed without a working folder.", vbExclamation, AppTitle)
        MousePointer = vbNormal
        Exit Sub
    Else
        GetPath = objVSSObject.LocalSpec
    End If
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Instantiate the selected version
    Set objVSSVersion = objVSSObject.Version(Val(Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " ")))))
    
    ' Get the file (checking for writeable copy)
    If GetAttr(GetPath) = vbArchive Then
        If Err <> 53 Then
            Response = MsgBox("File '" + GetPath + "' is writeable. Replace?", vbQuestion + vbYesNo, AppTitle)
        End If
    Else
        Response = vbYes
    End If
    If Response = vbYes Then
        objVSSVersion.Get Local:=GetPath, iFlags:=VSSFLAG_REPREPLACE
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
        
        Response = MsgBox("Unable to Get the selected version." + vbCrLf + Err.Description, vbExclamation, AppTitle)
   
   End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' View the file

Private Sub cmdView_Click()

    Dim objVSSVersion As VSSItem
    Dim lpBuffer As String * 100
    Dim TempPath As String
    Dim TempPathLength As Integer
    Dim Response As Long
    Dim RetVal As Long
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    Response = vbYes
    lpBuffer = String(100, Chr(0))

    ' Get path to Windows Temp Directory
    TempPathLength = GetTempPath(100, lpBuffer)
    TempPath = Mid(lpBuffer, 1, TempPathLength)
    
    ' Instantiate the selected version
    Set objVSSVersion = objVSSObject.Version(Val(Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " ")))))

    ' Set the "Get" path
    If Mid(TempPath, Len(TempPath), 1) = "\" Then
        TempPath = TempPath + objVSSVersion.Name
    Else
        TempPath = TempPath + "\" + objVSSVersion.Name
    End If
    
    ' View File based on users selection (OCX or OLE)
    If ViewMethod Then
    
        objVSSVersion.Get Local:=TempPath, iFlags:=VSSFLAG_REPREPLACE
        frmMain.Viewer1.ViewFile (TempPath)
        
    Else
    
        ' Instantiate the selected version
        Set objVSSVersion = objVSSObject.Version(Val(Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " ")))))
    
        ' File is binary
        If objVSSVersion.Binary = True Then
            Response = MsgBox("The file " + objVSSVersion.Name + " is binary. Continue?", vbYesNo + vbQuestion, AppTitle)
        End If
        If Response = vbYes Then
            
            ' Get the item to the Temp dir and display it
            objVSSVersion.Get Local:=TempPath, iFlags:=0
            RetVal = Shell("Notepad " + TempPath, vbNormalFocus)
    
        End If
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
        
        Response = MsgBox("Unable to view selected version." + vbCrLf + Err.Description, vbExclamation, AppTitle)
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal

End Sub

Private Sub cmdDiff_Click()

    Dim IsDifferent As Boolean
    Dim Response As Long
    Dim objVSSVersion1 As VSSItem
    Dim objVSSVersion2 As VSSItem
    Dim CheckOutFolder As String
    Dim TempPathLength As Long
    Dim lpFileData As String * 100
    Dim TempPath As String
    Dim GetDirectory As String
    Dim GetDirectoryV2 As String
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    lpFileData = String(100, Chr(0))
    Response = vbYes
    lpFileData = String(100, Chr(0))
    TempPath = GetTempPath(100, lpFileData)
    GetDirectory = Mid(lpFileData, 1, TempPath)
    
    ' Diffing VSS Copy against local copy
    If MultipleFilesSelected = False Then

        ' Instantiate the selected version
       Set objVSSVersion1 = objVSSObject.Version(Val(Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " ")))))
       
       ' Get the CheckOut Folder
       CheckOutFolder = objVSSVersion1.LocalSpec
       
       ' Diff using OLE
        If Not DiffMethod Then
        
            ' Check for differences
            IsDifferent = objVSSVersion1.IsDifferent(Local:=CheckOutFolder)
        
        Else
            
            ' Set Temp Dir Path for getting file
            If Mid(TempPath, Len(TempPath), 1) = "\" Then
                GetDirectory = GetDirectory + "SourceSafe version"
            Else
                GetDirectory = GetDirectory + "\SourceSafe version"
            End If
    
            ' Instantiate the selected item in the History window
            Set objVSSVersion1 = objVSSObject.Version(Val(Trim$(Mid(ListView1.SelectedItem.Key, InStr(ListView1.SelectedItem.Key, " ")))))
    
            ' Get the file
            objVSSVersion1.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
            frmMain.Diff1.DiffTwoFiles GetDirectory, CheckOutFolder, frmMain.hWnd
            
        End If
        
    ' Diffing two versions
    Else
        
        'Instantiate the two versions
        Set objVSSVersion1 = objVSSObject.Version(VersionNumberOne)
        Set objVSSVersion2 = objVSSObject.Version(VersionNumberTwo)
        
        ' Set Get Directory for version
        GetDirectory = GetDirectory + "SourceSafe version"
        GetDirectoryV2 = GetDirectory
             
        ' Diff using OLE
        If Not DiffMethod Then

            ' Get Version
            objVSSVersion2.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
            IsDifferent = objVSSVersion1.IsDifferent(Local:=CheckOutFolder)
             
         ' Diff using the OCX Control
         Else
            
             ' Set GetDirectory for versions
            GetDirectory = GetDirectory + Str(VersionNumberTwo)
            GetDirectoryV2 = GetDirectoryV2 + Str(VersionNumberOne)
            
            ' Get the files
            objVSSVersion1.Get Local:=GetDirectoryV2, iFlags:=VSSFLAG_REPREPLACE
            objVSSVersion2.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
            frmMain.Diff1.DiffTwoFiles GetDirectoryV2, GetDirectory, frmMain.hWnd
             
         End If
     End If
    
    ' If using OLE then continue...
    If Not DiffMethod Then Call ShowDiffInfo(IsDifferent, objVSSObject.Binary)
    
    'Check for errors
    If Err <> 0 Then
ErrHandler:

        Response = MsgBox("Unable to Diff objects." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize variables
    MultipleFilesSelected = False
    
End Sub

Private Sub ListView1_Click()
    
    ' Showing File History
    If FileHistory = True Then
        
        Dim VersionCount As Integer
        Dim Count As Integer
        
        ' Initialize variables
        Count = 0
        
        ' Iterate through Versions list
        For VersionCount = 1 To ListView1.ListItems.Count
        
            ' Version is selected
            If ListView1.ListItems(VersionCount).Selected = True Then
                Count = Count + 1
                
                ' Set command button's enabled state
                Select Case Count
                    
                    Case Is > 2
                        
                        cmdDiff.Enabled = False
                        cmdView.Enabled = False
                        cmdDetails.Enabled = False
                        MultipleFilesSelected = False
                        cmdGet.Enabled = False
                        Exit For
                        
                    Case 2
                        
                        MultipleFilesSelected = True
                        VersionNumberOne = Val(Trim$(Mid(ListView1.ListItems(VersionCount).Key, InStr(ListView1.SelectedItem.Key, " "))))
                        cmdView.Enabled = False
                        cmdDiff.Enabled = True
                        cmdDetails.Enabled = False
                        cmdGet.Enabled = False
                        
                    Case Else
                    
                        VersionNumberTwo = Val(Trim$(Mid(ListView1.ListItems(VersionCount).Key, InStr(ListView1.SelectedItem.Key, " "))))
                        MultipleFilesSelected = False
                        cmdDiff.Enabled = True
                        cmdView.Enabled = True
                        cmdDetails.Enabled = True
                        cmdGet.Enabled = True
                End Select
            End If
        Next
    End If
End Sub


