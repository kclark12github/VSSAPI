VERSION 5.00
Begin VB.Form frmNewDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New VSS Database"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "NewDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   345
      Left            =   3390
      TabIndex        =   1
      Top             =   345
      Width           =   1155
   End
   Begin VB.CheckBox chkNewFormat 
      Caption         =   "&Use new VSS 6.0 format"
      Height          =   300
      Left            =   165
      TabIndex        =   2
      Top             =   825
      Width           =   2970
   End
   Begin VB.TextBox txtDBPath 
      Height          =   285
      Left            =   165
      TabIndex        =   0
      Top             =   375
      Width           =   3060
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   983
      TabIndex        =   3
      Top             =   1470
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2543
      TabIndex        =   4
      Top             =   1470
      Width           =   1155
   End
   Begin VB.Label lblCreateIn 
      AutoSize        =   -1  'True
      Caption         =   "Create new database in:"
      Height          =   195
      Left            =   165
      TabIndex        =   5
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmNewDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()

    frmBrowseNewDB.Show 1

End Sub

Private Sub cmdCancel_Click()
    
    ' Close Form
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim Response As Long
    Dim VSSPath As String
    Dim Filehandle As Long
    Dim UsersText As String
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Get the VSS path
    VSSPath = GetVSSPath(objVSSDatabase.SrcSafeIni)
    If Right(VSSPath, 1) <> "\" Then VSSPath = VSSPath + "\"
    
    ' Verify target folder and create if needed
    If Dir(txtDBPath, vbDirectory) = "" Then
        MkDir txtDBPath
    End If
    
    ' Set Target Folder Structure
    If Right(txtDBPath, 1) <> "\" Then txtDBPath = txtDBPath + "\"
    
    If Dir(txtDBPath + "Data", vbDirectory) = "" Then
        MkDir txtDBPath + "Data"
    End If
    If Dir(txtDBPath + "Win32", vbDirectory) = "" Then
        MkDir txtDBPath + "Win32"
    End If
    If Dir(txtDBPath + "Template", vbDirectory) = "" Then
        MkDir txtDBPath + "Template"
    End If
    
    ' Copy EXE's and other support files
    FileCopy VSSPath + "Win32\MKSS.EXE", txtDBPath + "Win32\MKSS.EXE"
    FileCopy VSSPath + "Win32\DDCONV.EXE", txtDBPath + "Win32\DDCONV.EXE"
    FileCopy VSSPath + "Win32\DDUPD.EXE", txtDBPath + "Win32\DDUPD.EXE"
    FileCopy VSSPath + "Win32\SSUS.DLL", txtDBPath + "Win32\SSUS.DLL"
    FileCopy VSSPath + "Win32\SSAPI.DLL", txtDBPath + "Win32\SSAPI.DLL"
    FileCopy VSSPath + "Win32\ANALYZE.EXE", txtDBPath + "Win32\ANALYZE.EXE"
    FileCopy VSSPath + "Template\SrcSafe.ini", txtDBPath + "Template\SrcSafe.ini"
    FileCopy VSSPath + "Template\SSAdmin.ini", txtDBPath + "Template\SSAdmin.ini"
    
    ' Make target Database
    Shell (VSSPath + "WIN32\MKSS.EXE " + txtDBPath + "Data")
    Sleep 2000
    Shell (VSSPath + "WIN32\DDCONV.EXE " + txtDBPath + "Data")
    Sleep 2000
    If chkNewFormat.Value = 1 Then
        Shell (VSSPath + "WIN32\DDUPD.EXE " + txtDBPath + "Data")
    End If
    
    ' Make Target Directories
    If Dir(txtDBPath + "Users", vbDirectory) = "" Then
        MkDir txtDBPath + "Users"
    End If
    If Dir(txtDBPath + "Users\Admin", vbDirectory) = "" Then
        MkDir txtDBPath + "Users\Admin"
    End If
    If Dir(txtDBPath + "Users\Guest", vbDirectory) = "" Then
        MkDir txtDBPath + "Users\Guest"
    End If
    If Dir(txtDBPath + "Temp", vbDirectory) = "" Then
        MkDir txtDBPath + "Temp"
    End If

    ' Make target INI files
    FileCopy VSSPath + "Template\SrcSafe.ini", txtDBPath + "\SrcSafe.ini"
    FileCopy VSSPath + "Template\ssAdmin.ini", txtDBPath + "\Users\Admin\ssAdmin.ini"
    FileCopy VSSPath + "Users\Template.ini", txtDBPath + "\Users\Admin\ss.ini"
    FileCopy VSSPath + "Users\Template.ini", txtDBPath + "\Users\Guest\ss.ini"
    FileCopy VSSPath + "Users\Template.ini", txtDBPath + "Users\Template.ini"
    
    ' Make Users.txt
    If Dir(txtDBPath + "\Users.txt") = "" Then
        Filehandle = FreeFile
        UsersText = "admin=Users\admin\ss.ini" + vbCrLf + "guest=Users\guest\ss.ini" + vbCrLf
        Open txtDBPath + "Users.txt" For Output As #Filehandle
        Print #Filehandle, UsersText
        Close #Filehandle
    End If
                
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:
        If Err = 53 Then
            Response = MsgBox(" This command may only be run from a database that has a WIN32 directory containing" + vbCrLf + "MKSS.EXE, DDCONV.EXE and DDUPD.EXE. Please select an appropriate database and try again.", vbCritical, AppTitle)
            Unload Me
        ElseIf Err = 76 Then
            Response = MsgBox("Cannot find files required to create database. Possibly missing the local ...\VSS\Template folder." + vbCrLf, vbExclamation, AppTitle)
        Else
            Response = MsgBox("Cannot create database." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
        Err.Clear
    Else
        
        ' Close Form
        Me.SetFocus
        Unload Me
        MsgBox ("Database creation complete!")
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
            
End Sub

Private Sub Form_Load()
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub

Private Sub txtDBPath_Change()
    
    ' Enable OK Command Button if text has been entered
    If txtDBPath.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
End Sub
