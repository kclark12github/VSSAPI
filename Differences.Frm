VERSION 5.00
Begin VB.Form frmDifferences 
   Caption         =   "Differences"
   ClientHeight    =   5640
   ClientLeft      =   3600
   ClientTop       =   2355
   ClientWidth     =   7575
   Icon            =   "Differences.frx":0000
   LinkTopic       =   "Form19"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5640
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   6435
      TabIndex        =   2
      Top             =   5205
      Width           =   1050
   End
   Begin VB.TextBox txtVersion1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   315
      Width           =   3615
   End
   Begin VB.TextBox txtVersion2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3825
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   315
      Width           =   3615
   End
   Begin VB.Label lblLocalVersion 
      AutoSize        =   -1  'True
      Caption         =   "Local Version:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   45
      Width           =   1005
   End
   Begin VB.Label lblVSSVersion 
      AutoSize        =   -1  'True
      Caption         =   "SourceSafe Version:"
      Height          =   195
      Left            =   3825
      TabIndex        =   3
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "frmDifferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    ' Close the form
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim objVSSVersionOne As VSSItem
    Dim objVSSVersionTwo As VSSItem
    Dim FileOne As String
    Dim FileData As String
    Dim LineNumber As Integer
    Dim GetDirectory As String
    Dim lpFileData As String * 100
    Dim TempPath As String
    Dim Count As Integer
    Dim SelectedCount As Integer
    Dim VersionOne As Integer
    Dim VersionTwo As Integer
    Dim WorkingFolder As String
    Dim Response As Long
    
    ' Initialize variables
    LineNumber = 1
    lpFileData = String(100, Chr(0))
    TempPath = GetTempPath(100, lpFileData)
    GetDirectory = Mid(lpFileData, 1, TempPath)
    SelectedCount = 0
    VersionOne = 0
    VersionTwo = 0
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Calling from History Window
    If frmHistory.Visible = True Then
        
        ' Instantiate the selected item in the History window
        Set objVSSVersionOne = objVSSObject.Version(Val(Trim$(Mid(frmHistory.ListView1.SelectedItem.Key, InStr(frmHistory.ListView1.SelectedItem.Key, " ")))))
    
        ' Check to see if Diffing two versions
        For Count = 1 To frmHistory.ListView1.ListItems.Count
            If frmHistory.ListView1.ListItems(Count).Selected = True Then
                SelectedCount = SelectedCount + 1
                If VersionOne = 0 Then
                    VersionOne = Val(frmHistory.ListView1.ListItems(Count).Text)
                ElseIf VersionTwo = 0 And VersionOne <> 0 Then
                    VersionTwo = Val(frmHistory.ListView1.ListItems(Count).Text)
                End If
            End If
        Next Count
    
    ' Calling from frmMain
    Else
        
        ' Instantiate the selected item in the History window
        Set objVSSVersionOne = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
    End If
    
    ' Set Temp Dir Path for getting file
    If Mid(GetDirectory, Len(GetDirectory), 1) = "\" Then
        GetDirectory = GetDirectory + objVSSVersionOne.Name
    Else
        GetDirectory = GetDirectory + "\" + objVSSVersionOne.Name
    End If
    
    ' Set Form Caption
    frmDifferences.Caption = "Differences for file " + objVSSVersionOne.Name
    
    ' Diffing two versions
    If SelectedCount = 2 Then
    
        ' Instantiate first version
        Set objVSSVersionOne = objVSSObject.Version(VersionOne)

        ' Set the Left Pane Caption Caption
        frmDifferences.lblVSSVersion.Caption = "SourceSafe Version: " + Str(objVSSVersionOne.VersionNumber)
        
        ' Get the file
        objVSSVersionOne.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
        
        ' Read the file
        Open GetDirectory For Input As #1
        Line Input #1, FileData
        FileOne = Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        While Not EOF(1)
            Line Input #1, FileData
            LineNumber = LineNumber + 1
            FileOne = FileOne + vbCrLf + Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        Wend
        
        ' Populate the Left Pane and close the file
        txtVersion2.Text = FileOne
        Close (1)
        
        ' Delete the temp file
        SetAttr GetDirectory, vbNormal
        Kill GetDirectory
        
        ' Reset Line number
        LineNumber = 1
        
        ' Instantiate the selected item in the History window
        Set objVSSVersionTwo = objVSSObject.Version(VersionTwo)
        
        ' Set the Right Pane Caption Caption
        frmDifferences.lblLocalVersion.Caption = "SourceSafe Version: " + Str(objVSSVersionTwo.VersionNumber)
        
        ' Get the file
        objVSSVersionTwo.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
        
        ' Read the file
        Open GetDirectory For Input As #1
        Line Input #1, FileData
        FileOne = Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        While Not EOF(1)
            Line Input #1, FileData
            LineNumber = LineNumber + 1
            FileOne = FileOne + vbCrLf + Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        Wend
        
        ' Populate the Right Pane and close the file
        txtVersion1.Text = FileOne
        Close (1)
        
        ' Delete the temp file
        SetAttr GetDirectory, vbNormal
        Kill GetDirectory
    
    ' Diffing against a local copy
    Else
    
        ' Get location of local copy
        WorkingFolder = objVSSVersionOne.Parent.LocalSpec
        If Mid(WorkingFolder, Len(WorkingFolder), 1) = "\" Then
            WorkingFolder = WorkingFolder + objVSSVersionOne.Name
        Else
            WorkingFolder = WorkingFolder + "\" + objVSSVersionOne.Name
        End If
        
        ' Set Left Pane caption
        frmDifferences.lblLocalVersion.Caption = "Local Version: " + WorkingFolder
        
        ' Read file
        Open WorkingFolder For Input As #1
        Line Input #1, FileData
        FileOne = Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        While Not EOF(1)
            Line Input #1, FileData
            LineNumber = LineNumber + 1
            FileOne = FileOne + vbCrLf + Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        Wend
        
        ' Populate Left Pane and close file
        txtVersion1.Text = FileOne
        Close (1)
        
        ' Reset line number
        LineNumber = 1
        
        ' Set the Right Pane Caption Caption
        frmDifferences.lblVSSVersion.Caption = "SourceSafe Version: " + Str(objVSSVersionOne.VersionNumber)
        
        ' Get the VSS Copy file
        objVSSVersionOne.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
        
        ' Read the file
        Open GetDirectory For Input As #1
        Line Input #1, FileData
        FileOne = Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        While Not EOF(1)
            Line Input #1, FileData
            LineNumber = LineNumber + 1
            FileOne = FileOne + vbCrLf + Trim(Str(LineNumber)) + Space(5 - Len(Str(LineNumber))) + FileData
        Wend
        txtVersion2.Text = FileOne
        Close (1)
        
        ' Delete the temp file
        SetAttr GetDirectory, vbNormal
        Kill GetDirectory
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
           
        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Close (1)
        Err.Clear
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub
