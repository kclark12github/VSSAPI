VERSION 5.00
Begin VB.Form frmFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Format"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   Icon            =   "Format.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Frame frmFormat 
      Caption         =   "Select New Database Format"
      Height          =   960
      Left            =   105
      TabIndex        =   4
      Top             =   555
      Width           =   3420
      Begin VB.OptionButton optVSS6 
         Caption         =   "&New VSS 6.0 format (enhanced)"
         Height          =   270
         Left            =   210
         TabIndex        =   1
         Top             =   615
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optVSS5 
         Caption         =   "&VSS 5.0 format"
         Height          =   270
         Left            =   210
         TabIndex        =   0
         Top             =   285
         Width           =   3075
      End
   End
   Begin VB.Label lblCurrentFormat 
      AutoSize        =   -1  'True
      Caption         =   "CurrentFormat"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   195
      Width           =   990
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    ' Close form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim VSSPath As String
    Dim RetVal As Long
    Dim Response As Long
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Get the VSS path
    VSSPath = GetVSSPath(objVSSDatabase.SrcSafeIni)
    If Right(VSSPath, 1) <> "\" Then VSSPath = VSSPath + "\"

    ' Format database as requested, first warn user that we will shut
    ' down during conversion
    Response = MsgBox("This application will close while the conversion takes place. Continue?", vbYesNo + vbQuestion, AppTitle)
    If Response = vbYes Then
        
        ' Disconnect from the VSS Database Object
        Set objVSSDatabase = Nothing
        Set objVSSProject = Nothing
        Set objVSSVersion = Nothing
        Set objVSSCheckOut = Nothing
        Set objVSSObject = Nothing
    
        If optVSS6.Value = True Then
            RetVal = Shell(VSSPath + "WIN32\DDUPD.EXE " + VSSPath + "Data -redo", vbNormalFocus)
        Else
            RetVal = Shell(VSSPath + "WIN32\DDUPD.EXE " + VSSPath + "Data -undo", vbNormalFocus)
        End If
        End
    End If
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
    
        If Err = 53 Then
            Response = MsgBox(" This command may only be run from a database that has a WIN32 directory containing" + vbCrLf + "DDUPD.EXE. Please select an appropriate database and try again.", vbCritical, AppTitle)
            Unload Me
        Else
            Response = MsgBox("Cannot change database format." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
    
    Else
        
        ' Close Form
        Me.SetFocus
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    
    Dim VSSPath As String
    Dim DatabaseFormat As String
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Get and show current Database Format
    VSSPath = GetVSSPath(objVSSDatabase.SrcSafeIni)
    If Dir(VSSPath + "\Data\Labels", vbDirectory) = "" Then
        lblCurrentFormat.Caption = "Current format is VSS 5.0"
        optVSS6.Value = True
    Else
        lblCurrentFormat.Caption = "Current format is VSS 6.0 (enhanced)"
        optVSS5.Value = True
    End If
    
End Sub
