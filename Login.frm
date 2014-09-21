VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In To SourceSafe"
   ClientHeight    =   2355
   ClientLeft      =   4590
   ClientTop       =   5355
   ClientWidth     =   4260
   ClipControls    =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   4260
   Begin VB.CheckBox chkOpenDatabase 
      Caption         =   "Open this database next time I run SourceSafe"
      Height          =   315
      Left            =   225
      TabIndex        =   3
      Top             =   1485
      Value           =   1  'Checked
      Width           =   3870
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   1485
      TabIndex        =   5
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   1920
      Width           =   1275
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   135
      Width           =   2355
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   585
      Width           =   2355
   End
   Begin VB.TextBox txtSrcsafeini 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   990
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Path to SRCSAFE.INI:"
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   1050
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   645
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Password::"
      Height          =   195
      Left            =   -2520
      TabIndex        =   7
      Top             =   1125
      Width           =   780
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()

    ' This routine is called when the user attempts to log into
    ' SourceSafe. It verifies the Username, Password and
    ' SRCSAFE.INI. If any of these are invalid, the logon attempt
    ' will fail and the application will display an error
    ' message.
    
    Dim Response As Long
    Dim Filehandle As Integer
    Dim IniData As String
    Dim dbName As String
       
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Show splash screen
    If Not (NoSplashScreen) Then
        frmSplash.Show
        frmSplash.Refresh
        Sleep 1000
    End If
    
    ' Set on error routine
    On Error Resume Next
    
    ' Initialize the Username, password and srcsafe.ini variables
    UserName = UCase(Mid(frmLogon.txtUserName.Text, 1, 1)) + LCase(Mid(frmLogon.txtUserName.Text, 2))
    Password = frmLogon.txtPassword
    SrcSafeIni = frmLogon.txtSrcsafeini
    If UCase(Right(SrcSafeIni, 11)) <> "SRCSAFE.INI" Then
        If Right(SrcSafeIni, 1) = "\" Then
            SrcSafeIni = SrcSafeIni + "SRCSAFE.INI"
        Else
            SrcSafeIni = SrcSafeIni + "\SRCSAFE.INI"
        End If
    End If
    
    ' Set the Database object to nothing (in case
    ' user is logging in to a new database)
    Set objVSSDatabase = Nothing
        
    ' Attempt to log into SourceSafe
    objVSSDatabase.Open SrcSafeIni, UserName, Password
    
    ' If an error occured show appropriate dialog
    Select Case Err
        
        Case 0
            
            ' Save logon data if appropriate
            If chkOpenDatabase.Value = 1 Then
                dbName = objVSSDatabase.DatabaseName
                IniData = "[Username]" + vbCrLf + UserName + vbCrLf + vbCrLf + "[Database Path]" + vbCrLf + SrcSafeIni + vbCrLf + vbCrLf + "[Database Name]" + vbCrLf + dbName
                Filehandle = FreeFile
                Open AppPath + "srcsafeole.ini" For Output As #Filehandle
                Print #Filehandle, IniData
                Close #Filehandle
            End If
            
            ' If no error then open the main window
            Unload Me
            If frmMain.Visible Then
                frmMain.TreeView1.Nodes.Clear
                frmMain.ListView1.ListItems.Clear
                frmMain.Form_Load
            Else
                frmMain.Show
            End If
        Case Else
        
            ' Display error message
            Response = MsgBox("Error logging into SourceSafe!" + vbCrLf + Err.Description + ".", vbExclamation, AppTitle)
            Err.Clear
            
            ' If an error occured we need to clear the current database object
            ' to reattempt logon
            Set objVSSDatabase = Nothing
    End Select
                
    ' Unload splash screen
    Unload frmSplash
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

Private Sub cmdCancel_Click()

    ' Cancel Login and close application
    Unload Me
    
End Sub

Private Sub cmdBrowse_Click()

    ' Show Browse for SRCSAFE.INI dialog
    frmBrowse.Show 1
    
End Sub

Private Sub Form_Load()

    Dim Filehandle As Integer
    Dim IniData As String

    ' Set on error routine
    On Error Resume Next

    ' Center Form
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Set the AppPath variable to the applications path
    If Right(App.Path, 1) <> "\" Then
        AppPath = App.Path + "\"
    Else
        AppPath = App.Path
    End If
    
    ' Look for the INI file. If found load saved database info
    Filehandle = FreeFile
    Open AppPath + "srcsafeole.ini" For Input As #Filehandle
    If Err = 0 Then
        Do While Not EOF(Filehandle)
            Input #Filehandle, IniData
            Select Case IniData
                Case "[Username]"
                    Input #Filehandle, IniData
                    txtUserName.Text = IniData
                Case "[Database Path]"
                    Input #Filehandle, IniData
                    txtSrcsafeini.Text = IniData
                Case "[Database Name]"
                    Input #Filehandle, IniData
                    If Trim(IniData) <> "" Then
                        frmLogon.Caption = frmLogon.Caption + " - " + IniData
                    End If
            End Select
            chkOpenDatabase.Value = 1
        Loop
        Close #Filehandle
    Else
        
        ' No ini file found so clear error
        Err.Clear
    End If
End Sub

Private Sub txtUserName_Change()

    ' Call routine to check if there is text in the Username
    ' and SRCSAFE.INI Text boxes. Enable OK Command Button as
    ' appropriate
    cmdOK.Enabled = Checkdata
    
End Sub

Private Sub txtSrcsafeini_Change()
    
    ' Call routine to check if there is text in the Username
    ' and SRCSAFE.INI Text boxes. Enable OK Command Button as
    ' appropriate
    cmdOK.Enabled = Checkdata
    
End Sub

Public Function Checkdata() As Boolean
    
    ' Enable OK dialog if the Logon window contains text in the Username and SRCSAFE.INI Text boxes
    Checkdata = Len(txtUserName.Text) <> 0 And Len(txtSrcsafeini.Text) <> 0

End Function
