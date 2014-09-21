Attribute VB_Name = "Module1"
Option Explicit

' Global constant for Application title and Path
Global Const AppTitle = "Visual SourceSafe OLE Sample"
Global AppPath As String

' Used to store the VSS Username, password and SrcSafe.ini data
Global UserName As String
Global SrcSafeIni As String
Global Password As String

' Used as a reference to the VSS database
Global objVSSDatabase As New VSSDatabase

' Used to store a VSSItem Objects
Global objVSSProject As VSSItem
Global objVSSVersion As VSSVersion
Global objVSSCheckout As Object
Global objVSSObject As Object

' Used for user personal settings
Global WarnExit As Boolean
Global WarnDestroy As Boolean
Global WarnDelete As Boolean
Global WarnPurge As Boolean
Global WarnCheckOut As Boolean
Global WarnUndoCheckOut As Boolean
Global DoubleClickFile As String
Global ShowContext As Boolean

' Used to select an item in the list view when right click is issued
Global WhichButton As Integer

' Used to store a string representing the date/timestamp of a file
Global FileDate As String

' Used to indicate which pane is selected
Global Selected As String

' Used to store Flags
Global Flags As Long

' Constants used for dragging a share or branch operation
Global Dragging As Integer
Global Const EndDrag = 0
Global Const DragLeft = 1
Global Const DragRight = 2
Global Const DragRightBranch = 3
Global Const DragMove = 4

' Used to indicate if users other than Admin are logged into the Database
Global LoggedOn As Boolean

' Used when assigning project rights
Global CurrentUser As VSSUser

' Used to determine if splash screen should be displayed
Global NoSplashScreen As Boolean

' Used to tally the number of items in the current project
Global ItemCount As Integer

' Used to store backup folder for Analyze
Global BackupFolder As String

' Used for the ForceDir setting
Global ForceDirFlag As Long

' Used for the EOL setting
Global EOLFlag As Long

' Used for the CheckIn Unchanged Files setting
Global CheckInUnchangedFlag As Long

' Used to determine if we View\Edit\Diff files with OLE or OCX
Global ViewMethod As Boolean
Global DiffMethod As Boolean

' Used to flaf whether we are displaying file or project History
Global FileHistory As Boolean

' Declare Sleep API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' This function returns the Windows Temp directory
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' This function is passed a path including a file and returns the path minus the file
' For example it is passed C:\RIGEL\DOG.TXT and returns C:\RIGEL

Public Function GetDirPath(FilePath As String) As String
    
    Dim TestCharacter As String
    Dim Count As Integer
    
    While TestCharacter <> "\"
       TestCharacter = Mid(FilePath, Len(FilePath) - Count, 1)
       Count = Count + 1
    Wend
    TestCharacter = Mid(FilePath, 1, Len(FilePath) - Count)
    If Len(TestCharacter) = 2 Then TestCharacter = TestCharacter + "\"
    GetDirPath = TestCharacter
    
End Function

' This function is passed a path including a file name and returns
' the file name. For example it is passed C:\RIGEL\DOG.TXT it
' returns DOG.TXT

Public Function GetFileName(ByVal FilePath As String) As String
    
    While InStr(FilePath, "\") <> 0
       FilePath = Mid(FilePath, InStr(FilePath, "\") + 1)
    Wend
    GetFileName = FilePath
    
End Function

' This function is passed a path including a file name and returns
' the file name without the extension. For example it is passed
' C:\RIGEL\DOG.TXT it returns DOG

Public Function GetFileNameNoExt(ByVal FilePath As String) As String
    
    Dim Ext As Long
    
    While InStr(FilePath, "\") <> 0
       FilePath = Mid(FilePath, InStr(FilePath, "\") + 1)
    Wend
    
    Ext = InStr(FilePath, ".")
    If Ext <> 0 Then FilePath = Left(FilePath, Ext - 1)
    GetFileNameNoExt = FilePath
    
End Function

' This function is passed a VSS path including a file name and returns
' the file name. For example it is passed $/RIGEL/DOG.TXT and
' returns DOG.TXT

Public Function GetVSSFileName(ByVal FilePath As String) As String
    
    While InStr(FilePath, "/") <> 0
       FilePath = Mid(FilePath, InStr(FilePath, "/") + 1)
    Wend
    GetVSSFileName = FilePath
    
End Function

' This function is passed a path including a file and returns the path minus the file
' For example it is passed C:\RIGEL\DOG.TXT and returns C:\RIGEL

Public Function GetVSSPath(VSSPath As String) As String
    
    Dim TestCharacter As String
    Dim Count As Integer
    
    While TestCharacter <> "\"
       TestCharacter = Mid(VSSPath, Len(VSSPath) - Count, 1)
       Count = Count + 1
    Wend
    TestCharacter = Mid(VSSPath, 1, Len(VSSPath) - Count)
    If Len(TestCharacter) = 2 Then TestCharacter = TestCharacter + "\"
    GetVSSPath = TestCharacter
    
End Function

' Displays the working folder information on the main form

Public Sub DisplayWorkingFolder(objVSSProject As VSSItem)
    
    Dim WorkingFolder As String
    
    ' Set On Error routine
    On Error Resume Next
    
    ' Check for WorkingFolder
    WorkingFolder = objVSSProject.LocalSpec
    
    If WorkingFolder = "" Then
        Err.Clear
        frmMain.lblWorkingFolder.Caption = "No Working Folder"
    Else
        frmMain.lblWorkingFolder.Caption = "Working Folder: " + WorkingFolder
    End If
    
End Sub

' Sets the CheckOutFolder when diffing a VSS Version to a local copy

Public Function GetLocalCheckOutFolder(objVSSVersionObject As VSSItem) As String

    Dim CheckOutFolder As String
        
    ' File is checked out
    If objVSSVersionObject.IsCheckedOut = VSSFILE_CHECKEDOUT_ME Then
        For Each objVSSCheckout In objVSSVersionObject.Checkouts
            If objVSSCheckout.UserName = UserName Then
                CheckOutFolder = objVSSCheckout.LocalSpec
                Exit For
            End If
        Next
    
    ' File is not checked out
    Else
        CheckOutFolder = objVSSVersionObject.Parent.LocalSpec
    End If
    
    ' Set the working folder
    If Right(CheckOutFolder, 1) = "\" Then
        GetLocalCheckOutFolder = CheckOutFolder + objVSSVersionObject.Name
    Else
        GetLocalCheckOutFolder = CheckOutFolder + "\" + objVSSVersionObject.Name
    End If
    
End Function

' Displays info on two diffed files

Sub ShowDiffInfo(Different As Boolean, IsBinary As Boolean)

    Select Case IsBinary
    
        Dim Response As Long
            
        Case True
            
            If Different = False Then
                Response = MsgBox("Binary files are identical.", vbInformation, AppTitle)
            Else
                Response = MsgBox("Binary files differ.", vbInformation, AppTitle)
            End If
            
        Case False
            
            If Different = False Then
                Response = MsgBox("Text files are identical.", vbInformation, AppTitle)
            Else
                frmDifferences.Show 1
            End If
    End Select
End Sub

' Updates the results view window

Public Sub ShowResults()
    
    frmMain.txtResults.SetFocus
    SendKeys "^{END}", True
    If Selected = "Treeview1" Then
        frmMain.TreeView1.SetFocus
    Else
        frmMain.ListView1.SetFocus
    End If

End Sub
