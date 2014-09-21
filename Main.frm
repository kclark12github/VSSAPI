VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{5DC92A4A-C010-11D0-9A9A-00C04FC3066A}#1.0#0"; "DIFFME~1.OCX"
Begin VB.Form frmMain 
   Caption         =   "SourceSafe - Visual Basic Style"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   840
   ClientWidth     =   8895
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   8895
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   820
      ButtonWidth     =   741
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   21
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CreateProject"
            Object.ToolTipText     =   "Create Project"
            Object.Tag             =   ""
            ImageKey        =   "CreateProject"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AddFile"
            Object.ToolTipText     =   "Add Files"
            Object.Tag             =   ""
            ImageKey        =   "AddFiles"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LabelItem"
            Object.ToolTipText     =   "Label Version"
            Object.Tag             =   ""
            ImageKey        =   "LabelItem"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Files\Project"
            Object.Tag             =   ""
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GetFile"
            Object.ToolTipText     =   "Get Latest Version"
            Object.Tag             =   ""
            ImageKey        =   "GetLatestVersion"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CheckOut"
            Object.ToolTipText     =   "Check Out File"
            Object.Tag             =   ""
            ImageKey        =   "CheckOut"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CheckIn"
            Object.ToolTipText     =   "Check In File"
            Object.Tag             =   ""
            ImageKey        =   "CheckIn"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UndoCheckOut"
            Object.ToolTipText     =   "Undo Check Out"
            Object.Tag             =   ""
            ImageKey        =   "UndoCheckOut"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Share"
            Object.ToolTipText     =   "Share Files"
            Object.Tag             =   ""
            ImageKey        =   "Share"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Branch"
            Object.ToolTipText     =   "Branch Files"
            Object.Tag             =   ""
            ImageKey        =   "Branch"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ViewFile"
            Object.ToolTipText     =   "View File"
            Object.Tag             =   ""
            ImageKey        =   "ViewFile"
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EditFile"
            Object.ToolTipText     =   "Edit File"
            Object.Tag             =   ""
            ImageKey        =   "EditFile"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ShowDifferences"
            Object.ToolTipText     =   "File/Project Difference"
            Object.Tag             =   ""
            ImageKey        =   "ShowDifferences"
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Show Properties"
            Object.Tag             =   ""
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "History"
            Object.ToolTipText     =   "Show History"
            Object.Tag             =   ""
            ImageKey        =   "History"
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SetWorkingFolder"
            Object.ToolTipText     =   "Set Working Folder"
            Object.Tag             =   ""
            ImageKey        =   "SetWorkingFolder"
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh File List"
            Object.Tag             =   ""
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "About"
            Object.ToolTipText     =   "Shows the About Box"
            Object.Tag             =   ""
            ImageKey        =   "About"
         EndProperty
      EndProperty
   End
   Begin DIFFMERGECTLLib.Diff Diff1 
      Left            =   4410
      Top             =   525
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   317
      _StockProps     =   0
      DiffFormat      =   "visual"
      DiffWidth       =   512
      DiffContext     =   -1
      DiffIgnore      =   "WE"
      VisualDiffMax   =   0   'False
      VisualDiffModal =   0   'False
      VisualDiffRect  =   "100, 100, 600, 400"
   End
   Begin DIFFMERGECTLLib.Viewer Viewer1 
      Left            =   3780
      Top             =   510
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   344
      _StockProps     =   0
      SyntaxColoring  =   0   'False
      UseLineMarker   =   -1  'True
      ShowLineNumbers =   -1  'True
   End
   Begin VB.PictureBox PaneSeperatorBottom 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   60
      Left            =   75
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   8745
      TabIndex        =   9
      Top             =   4755
      Width           =   8745
   End
   Begin VB.TextBox txtResults 
      Height          =   945
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4830
      Width           =   8850
   End
   Begin VB.PictureBox PaneSeperator 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   6630
      Left            =   2085
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6630
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   840
      Width           =   75
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6210
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3915
      Left            =   2190
      TabIndex        =   2
      Top             =   840
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "NAME"
         Object.Tag             =   "NAME"
         Text            =   "Name:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "USER"
         Object.Tag             =   "USER"
         Text            =   "User:"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "DATETIME"
         Object.Tag             =   "DATETIME"
         Text            =   "Date-Time:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "FOLDER"
         Object.Tag             =   "FOLDER"
         Text            =   "Check Out Folder"
         Object.Width           =   2716
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   6030
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   423
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4921
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "USER"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Sort: Name"
            TextSave        =   "Sort: Name"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3900
      Left            =   15
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   6879
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   8055
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":014A
            Key             =   "NoDrop"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0464
            Key             =   "Shared"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":057E
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0690
            Key             =   "CheckedOutShared"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":07AA
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":08BC
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0AB0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0BC2
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0CBC
            Key             =   "Checked"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0DCE
            Key             =   "Labeled"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0ED0
            Key             =   "CheckedOutExclusive"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0FEA
            Key             =   "CheckedOutMulti"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":10FC
            Key             =   "CheckoutSharedMulti"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":120E
            Key             =   "NoDrop"
         EndProperty
      EndProperty
   End
   Begin DIFFMERGECTLLib.VisMerge VisMerge1 
      Left            =   5250
      Top             =   495
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   450
      _StockProps     =   0
      VisualMergeMax  =   0   'False
      VisualMergeRect =   "100, 100, 600, 400"
      VisualMerge     =   "Conflicts"
   End
   Begin VB.Line linePaneSepLeft 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   2055
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label lblAllProjects 
      Caption         =   "All Projects:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   570
      Width           =   825
   End
   Begin VB.Line LineProjectsBorderTop 
      BorderColor     =   &H00808080&
      X1              =   15
      X2              =   2055
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line lineProjectsBorder 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   765
      Y2              =   1035
   End
   Begin VB.Line PaneSepLeft 
      BorderColor     =   &H00FFFFFF&
      X1              =   2055
      X2              =   2055
      Y1              =   525
      Y2              =   795
   End
   Begin VB.Line lineWorkingFolderRight 
      BorderColor     =   &H00FFFFFF&
      X1              =   8835
      X2              =   8835
      Y1              =   525
      Y2              =   795
   End
   Begin VB.Line lineWorkingFolderBottom 
      BorderColor     =   &H00FFFFFF&
      X1              =   4875
      X2              =   8850
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line lineWorkingFolderLeft 
      BorderColor     =   &H00808080&
      X1              =   4875
      X2              =   4875
      Y1              =   525
      Y2              =   810
   End
   Begin VB.Line lineWorkingFolderBorder 
      BorderColor     =   &H00808080&
      X1              =   4905
      X2              =   8865
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line lineContentsRight 
      BorderColor     =   &H00FFFFFF&
      X1              =   4815
      X2              =   4815
      Y1              =   525
      Y2              =   810
   End
   Begin VB.Line lineContentsBorderBottom 
      BorderColor     =   &H00FFFFFF&
      X1              =   2175
      X2              =   4830
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line LinePaneSepRight 
      BorderColor     =   &H00808080&
      X1              =   2175
      X2              =   2175
      Y1              =   525
      Y2              =   795
   End
   Begin VB.Line lineContentsBorderTop 
      BorderColor     =   &H00808080&
      X1              =   2175
      X2              =   4815
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label lblContents 
      AutoSize        =   -1  'True
      Caption         =   "Contents of:"
      Height          =   195
      Left            =   2220
      TabIndex        =   7
      Top             =   570
      Width           =   855
   End
   Begin VB.Label lblWorkingFolder 
      AutoSize        =   -1  'True
      Caption         =   "Working Folder:"
      Height          =   195
      Left            =   4920
      TabIndex        =   6
      Top             =   570
      Width           =   1125
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6765
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1528
            Key             =   "GetLatestVersion"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":16B6
            Key             =   "CheckOut"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1844
            Key             =   "History"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":19D2
            Key             =   "CheckIn"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1B60
            Key             =   "UndoCheckOut"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1CEE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1E7C
            Key             =   "LabelItem"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":200A
            Key             =   "AddFiles"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2198
            Key             =   "CreateProject"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2326
            Key             =   "SetWorkingFolder"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":24B4
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2642
            Key             =   "ViewFile"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":27D0
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":295E
            Key             =   "Share"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2AEC
            Key             =   "ShowDifferences"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2C7A
            Key             =   "EditFile"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2E08
            Key             =   "Branch"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2F96
            Key             =   "About"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "O&pen SourceSafe Database..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "&Add Files..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCreateProject 
         Caption         =   "C&reate Project..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete..."
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Re&name..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuSetWorkingFolder 
         Caption         =   "Set &Working Folder..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLabel 
         Caption         =   "&Label..."
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuViewFile 
         Caption         =   "&View File..."
      End
      Begin VB.Menu mnuEditFile 
         Caption         =   "&Edit File..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu mnuInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSort 
         Caption         =   "S&ort"
         WindowList      =   -1  'True
         Begin VB.Menu mnuName 
            Caption         =   "&Name"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuUser 
            Caption         =   "&User"
         End
         Begin VB.Menu mnuDate 
            Caption         =   "&Date"
         End
         Begin VB.Menu mnuCheckOutFolder 
            Caption         =   "&Check Out Folder"
         End
      End
      Begin VB.Menu mnuRefreshFileList 
         Caption         =   "Refresh &File List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear results pane"
      End
   End
   Begin VB.Menu mnuSourceSafe 
      Caption         =   "&SourceSafe"
      Begin VB.Menu mnuGet 
         Caption         =   "&Get Latest Version..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuCheckOut 
         Caption         =   "Check &Out..."
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuCheckin 
         Caption         =   "Check &In..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuUndoCheckOut 
         Caption         =   "&Undo Check Out..."
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShare 
         Caption         =   "&Share..."
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "&Branch..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuHistory 
         Caption         =   "Show &History..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuShowDifferences 
         Caption         =   "Show &Differences..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change P&assword..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuRightPanePopupMenu 
      Caption         =   "RightPanePopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopView 
         Caption         =   "View..."
      End
      Begin VB.Menu mnuPopEdit 
         Caption         =   "Edit..."
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGet 
         Caption         =   "Get Latest Version..."
      End
      Begin VB.Menu mnuPopCheckOut 
         Caption         =   "Check Out..."
      End
      Begin VB.Menu mnuPopCheckIn 
         Caption         =   "Check In..."
      End
      Begin VB.Menu mnuPopUndoCheckOut 
         Caption         =   "Undo Check Out..."
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopHistory 
         Caption         =   "Show History..."
      End
      Begin VB.Menu mnuPopShowDifferences 
         Caption         =   "Show Differences..."
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuPopRename 
         Caption         =   "Rename..."
      End
      Begin VB.Menu mnuPopProperties 
         Caption         =   "Properties..."
      End
   End
   Begin VB.Menu mnuLeftPanePopupMenu 
      Caption         =   "LeftPanePopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCreateProject 
         Caption         =   "Create Project..."
      End
      Begin VB.Menu mnuPopSetWorkingFolder 
         Caption         =   "Set Working Folder..."
      End
      Begin VB.Menu mnuPopLabel 
         Caption         =   "Label..."
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGetProject 
         Caption         =   "Get Latest Version..."
      End
      Begin VB.Menu mnuPopCheckOutProject 
         Caption         =   "Check Out..."
      End
      Begin VB.Menu mnuPopCheckInProject 
         Caption         =   "Check In..."
      End
      Begin VB.Menu mnuPopUndoCheckOutProject 
         Caption         =   "Undo Check Out..."
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopShare 
         Caption         =   "Share..."
      End
      Begin VB.Menu mnuPopProjectHistory 
         Caption         =   "Show History..."
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDeleteProj 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuPopRenameProj 
         Caption         =   "Rename..."
      End
      Begin VB.Menu mnuPopPropertiesProject 
         Caption         =   "Properties..."
      End
   End
   Begin VB.Menu mnuLeftPaneDragMenu 
      Caption         =   "LeftPaneDragMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDragandShare 
         Caption         =   "Share"
      End
      Begin VB.Menu mnuDragandShareBranch 
         Caption         =   "Share and Branch"
      End
      Begin VB.Menu mnuDragandMove 
         Caption         =   "Move"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelShareDrag 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    
    Call EndShareDrag(Source, X, Y)
    
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    If TypeOf Source Is ListView And Dragging <> EndDrag Then ListView1.DragIcon = ImageList3.ListImages(1).Picture

End Sub

' This routine initializes/populates the main window
' and sets the current VSSItem to the root project ($/).

Public Sub Form_Load()

    Dim Nodex As Node
    Dim Count As Integer
    Dim ProjectToMoveTo As String
    Dim ProjectToMoveToPath As String
    Dim objMoveTo As VSSItem
    Dim SSIniPath As String
    Dim Filehandle As Long
    Dim IniData As String
    Dim Response As Long
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Center Window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    ' Initialize variables
    ProjectToMoveTo = objVSSDatabase.CurrentProject
    ProjectToMoveToPath = ProjectToMoveTo
    ItemCount = 0
    TreeView1.ImageList = ImageList1
    txtResults.Text = ""
    ViewMethod = True
    DiffMethod = True
    ShowContext = False
    Dragging = EndDrag
    
    ' Disable commands as appropriate
    Toolbar1.Buttons.Item("ViewFile").Enabled = False
    Toolbar1.Buttons.Item("EditFile").Enabled = False
    mnuViewFile.Enabled = False
    mnuEditFile.Enabled = False
    
    ' Initialize users personal settings. These default to True. Later, we
    ' read the Users SS.INI and change these if any have been customized.
    WarnDelete = True
    WarnExit = True
    WarnDestroy = True
    WarnPurge = True
    WarnCheckOut = True
    WarnUndoCheckOut = True
    DoubleClickFile = "Ask"
    Diff1.DiffFormat = "visual"
    Diff1.DiffIgnore = "we"
    frmMain.VisMerge1.VisualMerge = "Conflicts"
    
    ' Open Users SS.INI and read personal settings
    SSIniPath = GetDirPath(objVSSDatabase.SrcSafeIni) + "\Users\" + objVSSDatabase.UserName + "\ss.ini"
    
    Filehandle = FreeFile
    Open SSIniPath For Input As #Filehandle
    Do While Not EOF(Filehandle)
        Input #Filehandle, IniData
        
        If InStr(IniData, "Warn_Exit") <> 0 Then
            
            If Right(IniData, 2) = "es" Then
                WarnExit = True
            Else
                WarnExit = False
            End If
            
        End If
        If InStr(IniData, "Warn_Remove") <> 0 Then
            
            If Right(IniData, 2) = "es" Then
                WarnDelete = True
            Else
                WarnDelete = False
            End If
                
        End If
        If InStr(IniData, "Warn_Destroy ") <> 0 Then
            
            If Right(IniData, 2) = "es" Then
                WarnDestroy = True
            Else
                WarnDestroy = False
            End If
                
        End If
        If InStr(IniData, "Warn_Purge") <> 0 Then
            
            If Right(IniData, 2) = "es" Then
                WarnPurge = True
            Else
                WarnPurge = False
            End If
            
        End If
        If InStr(IniData, "Warn_Multiple_Checkout") <> 0 Then
        
            If Right(IniData, 2) = "es" Then
                WarnCheckOut = True
            Else
                WarnCheckOut = False
            End If
            
        End If
        If InStr(IniData, "Warn_Uncheckout") <> 0 Then
        
            If Right(IniData, 2) = "es" Then
                WarnUndoCheckOut = True
            Else
                WarnUndoCheckOut = False
            End If
            
        End If
        If InStr(IniData, "DoubleClick_File") <> 0 Then
        
            DoubleClickFile = Trim(Right(IniData, Len(IniData) - InStr(IniData, "=")))
            
        End If
        If InStr(IniData, "Diff_Format") <> 0 Then
            
            frmMain.Diff1.DiffFormat = Trim(Right(IniData, Len(IniData) - InStr(IniData, "=")))
        
        End If
        If InStr(IniData, "Diff_Ignore") <> 0 Then
        
            frmMain.Diff1.DiffIgnore = Trim(Right(IniData, Len(IniData) - InStr(IniData, "=")))
        
        End If
        If InStr(IniData, "Diff_Context") <> 0 Then
            
            ShowContext = True
            frmMain.Diff1.DiffContext = Trim(Right(IniData, Len(IniData) - InStr(IniData, "=")))
        
        End If
        If InStr(IniData, "Visual_Merge") <> 0 Then
            
            ShowContext = True
            frmMain.VisMerge1.VisualMerge = Trim(Right(IniData, Len(IniData) - InStr(IniData, "=")))
        
        End If
        
    Loop
    Close #Filehandle
    
    ' Set Force Dir Variable (currently not accessable from OLE Interface)
    ' so we set it hear
    ForceDirFlag = VSSFLAG_FORCEDIRYES
    
    ' Set EOL flag for text files currently not accessable from OLE Interface)
    ' so we set it hear
    EOLFlag = VSSFLAG_EOLCRLF
    
    ' Set then Checkin Unchanged files flag to UndoCheckOut
    CheckInUnchangedFlag = VSSFLAG_UPDUNCH
    
    ' Create VSS Database object and set current item to $/ (root project)
    Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
    
    ' Add Root Project to Treeview control
    TreeView1.LineStyle = tvwRootLines
    Set Nodex = TreeView1.Nodes.Add(, , "$", "$/", "Closed")
    TreeView1.Nodes(1).Selected = True
    TreeView1.Nodes(1).Expanded = True
    
    ' Populate the Project and File List
    Call PopulateMain(objVSSProject)
    
    ' Moved to the user's current project
    Set objMoveTo = objVSSDatabase.VSSItem(ProjectToMoveTo, False)
    If InStr(3, ProjectToMoveTo, "/") <> 0 Then
        ProjectToMoveTo = Left(ProjectToMoveToPath, InStr(3, ProjectToMoveToPath, "/") - 1)
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
    End If
    While InStr(Len(ProjectToMoveTo) + 2, ProjectToMoveToPath, "/") <> 0
        ProjectToMoveTo = Left(ProjectToMoveToPath, InStr(Len(ProjectToMoveTo) + 2, ProjectToMoveToPath, "/") - 1)
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
        frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
    Wend
    
    ProjectToMoveTo = ProjectToMoveToPath
    If ProjectToMoveTo = "$/" Then ProjectToMoveTo = "$"
    frmMain.TreeView1.Nodes(ProjectToMoveTo).Selected = True
    frmMain.TreeView1.Nodes(ProjectToMoveTo).Expanded = True
    
    ' Populate the Project and File List
    If ProjectToMoveTo <> "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem(ProjectToMoveTo, False)
        ListView1.ListItems.Clear
        Call PopulateMain(objVSSProject)
    End If
    
    ' Set glyphs for all projects to the closed icon
    For Each Nodex In TreeView1.Nodes
        If Nodex.Image = "Open" Then Nodex.Image = "Closed"
    Next
    
    ' Set glyph for selected project to open icon
    TreeView1.SelectedItem.Image = "Open"
    
    ' Show caption
    frmMain.Caption = AppTitle
    
    ' Initialize Staus Bar
    StatusBar1.Panels(2).Text = UserName
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
    ' Check For Errors
    If Err <> 0 Then
    
ErrHandler:

        Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Err.Clear
        
    End If
    
End Sub

' This routine resizes various controls when the user
' resizes the Main Form

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If frmMain.WindowState <> 1 Then
        TreeView1.Height = frmMain.Height - 3060
        ListView1.Height = frmMain.Height - 3060
        PaneSeperator.Height = TreeView1.Height
        Call ResizeHeader
        txtResults.Width = frmMain.Width - 110
        txtResults.Top = ListView1.Top + ListView1.Height + 80
        txtResults.Height = frmMain.Height - StatusBar1.Height - txtResults.Top - Toolbar1.Height - 200
        PaneSeperatorBottom.Top = txtResults.Top - 80
        PaneSeperatorBottom.Width = txtResults.Width
    End If
    
    If frmMain.Width < 8000 Then frmMain.Width = 8000
    If frmMain.Height < 3000 Then frmMain.Height = 3000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim Response As Long
    
    ' Check if user really wants to quit
    If WarnExit Then
        Response = MsgBox("This will end your " + AppTitle + " session.", vbQuestion + vbOKCancel, AppTitle)
    Else
        Response = vbOK
    End If
    If Response = vbOK Then
        
        ' Close all objects
        Set objVSSDatabase = Nothing
        Set objVSSProject = Nothing
        Set objVSSVersion = Nothing
        Set objVSSCheckout = Nothing
        Set objVSSObject = Nothing
        End
    Else
        Cancel = True
    End If
    
End Sub

Private Sub lblAllProjects_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)

End Sub

Private Sub lblAllProjects_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    If TypeOf Source Is ListView And Dragging <> EndDrag Then ListView1.DragIcon = ImageList3.ListImages(1).Picture

End Sub

' Resize the Contents label

Private Sub lblContents_Change()

    lblContents.Width = lineContentsBorderTop.X2 - lineContentsBorderTop.X1 - 100
    
End Sub

Private Sub lblContents_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)

End Sub

' Resize the Working Folder caption

Private Sub lblWorkingFolder_Change()

    lblWorkingFolder.Width = lineWorkingFolderBorder.X2 - lineWorkingFolderBorder.X1 - 100

End Sub

Private Sub lblWorkingFolder_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)
    
End Sub

Private Sub ListView1_Click()

    Dim Count As Integer
    Dim FileCount As Integer
    
    ' Disable view and edit if no files in list
    If ListView1.ListItems.Count = 0 Then
        Call EnableFileControls(False)
        Exit Sub
    End If
    
    ' Loop through file list looking for selected files
    Count = 0
    For FileCount = 1 To ListView1.ListItems.Count
              
        If ListView1.ListItems(FileCount).Selected = True Then
            
            ' Tally total files selected
            Count = Count + 1
            If Count > 1 Then
                ' Disable view and edit if multiple files selected
                Call EnableFileControls(False)
                Exit Sub
            End If
        End If
    Next
    Call EnableFileControls(True)

End Sub

' This routine is called when the user clicks on a column header
' of the listview control. It is used to sort the control based
' on the column selected.
    
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)

    Select Case ColumnHeader.Index
        
        ' Sort by Filename
        Case 1
            
            If ListView1.SortKey <> 0 Then
                Call SortByName
            Else
                Call InvertSortOrder
            End If
        
        ' Sort by Username
        Case 2
            
            If ListView1.SortKey <> 1 Then
                Call SortByUser
            Else
                Call InvertSortOrder
            End If
        
        ' Sort by Date
        Case 3
            
            If ListView1.SortKey <> 2 Then
                Call SortByDate
            Else
                Call InvertSortOrder
            End If
        
        ' Sort by CheckOut Folder
        Case 4
        
            If ListView1.SortKey <> 3 Then
                Call SortByFolder
            Else
                Call InvertSortOrder
            End If
            
    End Select
    
End Sub

Private Sub ListView1_DblClick()

    ' Set On Error Routine
    On Error Resume Next

    ' Disable view and edit if no files in list
    If ListView1.ListItems.Count = 0 Then
        Call EnableFileControls(False)
        Exit Sub
    End If

    Select Case DoubleClickFile
    
        Case "Edit File"
        
            mnuEditFile_Click
        
        Case "View File"
        
            mnuViewFile_Click
            
        Case "Ask"

            frmAsk.Caption = "File " + ListView1.SelectedItem.Key
            If Err = 0 Then
                frmAsk.Show 1
            Else
                Call EnableFileControls(False)
                Err.Clear
            End If
            
    End Select

End Sub

' This routine is designed to provide the look and feel of
' a split window between the ListView Control and the TreeView
' control. There may be better ways to do this but this is what
' I came up with.

Private Sub ListView1_DragDrop(Source As Control, X As Single, Y As Single)
    
    If TypeOf Source Is PictureBox And Source.Name = "PaneSeperator" Then
        Source.Left = X + TreeView1.Width
        TreeView1.Width = X + TreeView1.Width
        Call ResizeHeader
    ElseIf TypeOf Source Is PictureBox And Source.Name = "PaneSeperatorBottom" Then
        Source.Top = ListView1.Top + Y
        ListView1.Height = Y
        TreeView1.Height = Y
        txtResults.Top = ListView1.Top + ListView1.Height + 80
        txtResults.Height = StatusBar1.Top - txtResults.Top
    ElseIf TypeOf Source Is ListView Then
        Call EndShareDrag(Source, X, Y)
    End If
    
End Sub

Private Sub ListView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    If TypeOf Source Is ListView And Dragging <> EndDrag Then ListView1.DragIcon = ImageList3.ListImages(1).Picture

End Sub

' File List has focus

Private Sub ListView1_GotFocus()

    Selected = "Listview1"
    mnuBranch.Enabled = True
    Toolbar1.Buttons.Item("Branch").Enabled = True
    
End Sub

' Set the mouse button used

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    WhichButton = Button
    
End Sub

' Determine if user is in drag drop mode

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If WhichButton = vbLeftButton Then
        Dragging = DragLeft
        ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
        ListView1.Drag
    ElseIf WhichButton = vbRightButton Then
        Dragging = DragRight
        ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
        ListView1.Drag
    End If

End Sub

' Show popup menu if appropriate

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If WhichButton = vbRightButton Then
        frmMain.PopupMenu frmMain.mnuRightPanePopupMenu
    End If
    WhichButton = -99
    
End Sub

Private Sub mnuAbout_Click()

    Call About

End Sub

Private Sub mnuAddFile_Click()

    Call AddFile
    
End Sub

Private Sub mnuBranch_Click()

    Call Branch(ListView1.SelectedItem.Key)

End Sub

' End the drag drop operation

Private Sub mnuCancelShareDrag_Click()
    
    Dragging = EndDrag
    
End Sub

Private Sub mnuChangePassword_Click()

    Call ChangePassword

End Sub

Private Sub mnuClear_Click()

    txtResults.Text = ""

End Sub

Private Sub mnuCreateProject_Click()

    Call CreateProject
    
End Sub

Private Sub mnuCheckout_Click()

    Call CheckOut
    
End Sub

Private Sub mnuDate_Click()

    Call SortByDate
    
End Sub

Private Sub mnuDelete_Click()

    Call DeleteItem
    
End Sub

Private Sub mnuDragandMove_Click()
    
    ' Moving a file
    Dragging = DragMove

End Sub

Private Sub mnuDragandShare_Click()

    ' End the drag drop operation
    Dragging = DragLeft
    
End Sub

Private Sub mnuDragandShareBranch_Click()
    
    ' End the drag drop operation
    Dragging = DragRightBranch
    
End Sub

Public Sub mnuEditFile_Click()
    
    ' Edit File based on users selection (OCX or OLE)
    If Not ViewMethod Then
        Call EditFile
    Else
        Call EditFileOCX
    End If

End Sub

Private Sub mnuEditFileOCX_Click()

    Call EditFileOCX
    
End Sub

Private Sub mnuGet_Click()

    Call GetFile
    
End Sub

Private Sub mnuHistory_Click()
    
    Call ShowHistory
    
End Sub

Private Sub mnuInvertSelection_Click()

    Call InvertSelection
    
End Sub

Private Sub mnuLabel_Click()

    Call LabelItem
    
End Sub

Private Sub mnuMove_Click()

    Call MoveItem

End Sub

Private Sub mnuName_Click()

    Call SortByName
    
End Sub

Private Sub mnuCheckOutFolder_Click()

    Call SortByFolder
    
End Sub

Private Sub mnuOpen_Click()
    
    NoSplashScreen = True
    frmLogon.Show 1

End Sub

Private Sub mnuOptions_Click()

    Call Options

End Sub

Private Sub mnuPopCheckIn_Click()

    Call CheckIn
    
End Sub

Private Sub mnuPopCheckInProject_Click()

    Call CheckIn
    
End Sub

Private Sub mnuPopCheckOut_Click()
    
    Call CheckOut
    
End Sub

Private Sub mnuPopCheckOutProject_Click()

    Call CheckOut
    
End Sub

Private Sub mnuPopCreateProject_Click()

    Call CreateProject
    
End Sub

Private Sub mnuPopDelete_Click()

    Call DeleteItem
    
End Sub

Private Sub mnuPopDeleteProj_Click()

    Call DeleteItem
    
End Sub

Private Sub mnuPopEdit_Click()

    ' Edit File based on users selection (OCX or OLE)
    If Not ViewMethod Then
        Call EditFile
    Else
        Call EditFileOCX
    End If
    
End Sub

Private Sub mnuPopEditOCX_Click()

    Call EditFileOCX

End Sub

Private Sub mnuPopGet_Click()

    Call GetFile
    
End Sub

Private Sub mnuPopGetProject_Click()

    Call GetFile
    
End Sub

Private Sub mnuPopHistory_Click()
    
    Call ShowHistory
    
End Sub

Private Sub mnuPopLabel_Click()

    Call LabelItem
    
End Sub

Private Sub mnuPopProjectHistory_Click()
    
    Call ShowHistory
    
End Sub

Private Sub mnuPopProperties_Click()

    Call Properties
    
End Sub

Private Sub mnuPopPropertiesProject_Click()

    Call Properties
    
End Sub

Private Sub mnuPopRename_Click()

    Call Rename
    
End Sub

Private Sub mnuPopRenameProj_Click()

    Call Rename
    
End Sub

Private Sub mnuPopSetWorkingFolder_Click()

   Call SetWorkingFolder
   
End Sub

Private Sub mnuPopShare_Click()

   Call Share
   
End Sub

Private Sub mnuPopShowDifferences_Click()

    Call ShowDifferences

End Sub

Private Sub mnuPopUndoCheckOut_Click()

    Call UndoCheckOut
    
End Sub

Private Sub mnuPopUndoCheckOutProject_Click()

    Call UndoCheckOut
    
End Sub

Private Sub mnuPopView_Click()

    ' View File based on users selection (OCX or OLE)
    If Not ViewMethod Then
        Call ViewFile(ListView1.SelectedItem.Key)
    Else
        Call ViewFileOCX
    End If

End Sub

Private Sub mnuProperties_Click()

    Call Properties
    
End Sub

Private Sub mnuRefreshFileList_Click()

    Call RefreshFileList
    
End Sub

Private Sub mnuRename_Click()

    Call Rename
    
End Sub

Private Sub mnuShare_Click()

   Call Share
   
End Sub

Private Sub mnuShowDifferences_Click()

    Call ShowDifferences

End Sub

Private Sub mnuUndoCheckOut_Click()

    Call UndoCheckOut
    
End Sub

Private Sub mnuUser_Click()

    Call SortByUser
    
End Sub

Public Sub mnuViewFile_Click()

    ' View File based on users selection (OCX or OLE)
    If Not ViewMethod Then
        Call ViewFile(ListView1.SelectedItem.Key)
    Else
        Call ViewFileOCX
    End If
    
End Sub

Private Sub mnuSetWorkingFolder_Click()

   Call SetWorkingFolder
   
End Sub

Private Sub mnuCheckIn_Click()

    Call CheckIn
    
End Sub

Private Sub mnuEXIT_Click()

    Unload Me
    
End Sub

' This routine uses the Font Common Dialog to set the
' font for the TreeView and ListView controls

Private Sub mnuFont_Click()

    Dim objFont As New StdFont
    Dim Response As Long
    
    ' Set on Error routine
    On Error GoTo ErrHandler
    
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    objFont.Bold = CommonDialog1.FontBold
    objFont.Name = CommonDialog1.FontName
    objFont.Size = CommonDialog1.FontSize
    TreeView1.Font.Name = objFont.Name
    TreeView1.Font.Bold = objFont.Bold
    TreeView1.Font.Size = objFont.Size
    ListView1.Font.Name = objFont.Name
    ListView1.Font.Bold = objFont.Bold
    ListView1.Font.Size = objFont.Size
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:

        Response = MsgBox("Cannot set font as requested." + Err.Description, vbExclamation, AppTitle)
    End If
    
End Sub

' This routine is called by the Select All menu option. It selects
' all the files in the current project

Private Sub mnuSelectAll_Click()
    
    Dim Count As Integer
    
    For Count = 1 To ListView1.ListItems.Count
        ListView1.ListItems(Count).Selected = True
    Next
    ListView1.SetFocus
    
End Sub

Private Sub PaneSeperator_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)

End Sub

Private Sub PaneSeperatorBottom_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)
    
End Sub

Private Sub PaneSeperatorBottom_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    If TypeOf Source Is ListView And Dragging <> EndDrag Then ListView1.DragIcon = ImageList3.ListImages(1).Picture

End Sub

Private Sub StatusBar1_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)
    
End Sub

' Select appropriate routine based on the Toolbar selection

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)

    Select Case Button.Key
        Case "CheckOut"
            Call CheckOut
        Case "CheckIn"
            Call CheckIn
        Case "GetFile"
            Call GetFile
        Case "UndoCheckOut"
            Call UndoCheckOut
        Case "History"
            Call ShowHistory
        Case "Delete"
            Call DeleteItem
        Case "LabelItem"
            Call LabelItem
        Case "AddFile"
            Call AddFile
        Case "CreateProject"
            Call CreateProject
        Case "SetWorkingFolder"
            Call SetWorkingFolder
        Case "Properties"
            Call Properties
        Case "EditFile"
            Call EditFile
        Case "ViewFile"
            Call ViewFile(ListView1.SelectedItem.Key)
        Case "Refresh"
            Call RefreshFileList
        Case "Share"
            Call Share
        Case "Branch"
            Call Branch(ListView1.SelectedItem.Key)
        Case "ShowDifferences"
            Call ShowDifferences
        Case "About"
            Call About
    End Select
    
End Sub

Private Sub Toolbar1_DragDrop(Source As Control, X As Single, Y As Single)

    Call EndShareDrag(Source, X, Y)

End Sub

' This routine is called when the user clicks on a project in the TreeView
' control. It will set focus to the project, display it's contents (files)
' in the ListView control

Private Sub TreeView1_Click()
    
    Dim Nodex As Node
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working"
    
    ' Disable commands as appropriate
    Toolbar1.Buttons.Item("ViewFile").Enabled = False
    Toolbar1.Buttons.Item("EditFile").Enabled = False
    mnuViewFile.Enabled = False
    mnuEditFile.Enabled = False
    
    ' Set glyphs for all projects to the closed icon
    For Each Nodex In TreeView1.Nodes
        If Nodex.Image = "Open" Then Nodex.Image = "Closed"
    Next
    
    ' Set glyph for selected project to open icon
    TreeView1.SelectedItem.Image = "Open"
    
    ' Clear the ListView control of all files
    ListView1.ListItems.Clear
    
    ' Check if user has selected the root project and set the VSSItem as appropriate
    If TreeView1.SelectedItem.Key <> "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        lblContents.Caption = "Contents of: " + TreeView1.SelectedItem.Key
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        lblContents.Caption = "Contents of: $/"
    End If
    Call PopulateMain(objVSSProject)
    
    ' Reset mousepointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
End Sub

Private Sub TreeView1_Collapse(ByVal Node As Node)

    Node.Image = "Closed"
    
End Sub

' This routine is designed to provide the look and feel of a split window
' between the ListView Control and the TreeView control. There may be better
' ways to do this but this is what I came up with...

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)

    Dim Count As Integer
    Dim objFileListItem As ListItem
    Dim objVSSFile As VSSItem
    Dim Response As Long
     
    If TypeOf Source Is PictureBox And Source.Name = "PaneSeperator" Then
        
        Source.Left = X
        TreeView1.Width = Source.Left
        ListView1.Left = TreeView1.Width + PaneSeperator.Width
        ListView1.Width = frmMain.Width - TreeView1.Width + PaneSeperator.Width
        Call ResizeHeader
    
    ElseIf TypeOf Source Is PictureBox And Source.Name = "PaneSeperatorBottom" Then
       
        Source.Top = ListView1.Top + Y
        ListView1.Height = Y
        TreeView1.Height = Y
        txtResults.Top = ListView1.Top + ListView1.Height + 80
        txtResults.Height = StatusBar1.Top - txtResults.Top
        
    ElseIf TypeOf Source Is ListView Then
        
        ' User is sharing a file through drag and drop
        If Dragging = DragLeft Or Dragging = DragRight Or Dragging = DragMove Then
        
            ' Set On Error routine
            On Error Resume Next
            Dim ProjectTest As String
            ProjectTest = TreeView1.DropHighlight.Key
            
            ' No project selected in the Treeview control
            If Err = 91 Then
                Err.Clear
            Else
                       
                ' Set On Error routine
                On Error GoTo ErrHandler
                
                ' User is Right click dragging
                If Dragging = DragRight Then
                    
                    frmMain.PopupMenu frmMain.mnuLeftPaneDragMenu

                End If
                
                ' Check if user has canceled the drag
                If Dragging <> EndDrag Then
                
                    ' Instantiate the highlighted project
                    If frmMain.TreeView1.DropHighlight.Key = "$" Then
                        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
                    Else
                        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.DropHighlight.Key, False)
                    End If
                    
                    ' Iterate through the file list
                    For Count = 1 To ListView1.ListItems.Count
                        
                        Set objFileListItem = ListView1.ListItems(Count)
                    
                        ' File is selected
                        If objFileListItem.Selected = True Then
                        
                            ' Instantiate the file item
                            Set objVSSFile = objVSSDatabase.VSSItem(objFileListItem.Key, False)
                            
                            ' The move method currently does not support moving files. This
                            ' line will generate an error
                            If Dragging = DragMove Then
                                objVSSFile.Move objVSSProject
                            Else
                            
                                ' Share the file and update the main form's GUI
                                objVSSProject.Share pIItem:=objVSSFile, Comment:="", iFlags:=0
                                
                                ' Branch if required
                                If Dragging = DragRightBranch Then
                                    Set objVSSFile = objVSSDatabase.VSSItem(objFileListItem.Key, False)
                                    objVSSFile.Branch Comment:="", iFlags:=0
                                End If
                            End If
                        End If
ShareNext:
        
                    Next
            
                    ' Check for errors
                    If Err <> 0 Then
ErrHandler:
                            
                        Response = MsgBox("Unable to Share/Branch/Move file." + vbCrLf + Err.Description, vbExclamation, AppTitle)
                        Resume ShareNext
                        Err.Clear
                    End If
                End If
             End If
            
            ' End the drag operation
            Call EndShareDrag(Source, X, Y)
            Call RefreshFileList
        End If
    End If
    
End Sub

Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    If Dragging = DragLeft Or Dragging = DragRight Then Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    
    ' Set On Error routine
    On Error Resume Next
    Dim ProjectTest As String
    ProjectTest = TreeView1.DropHighlight.Key
    
    ' No project selected in the Treeview control
    If Err = 91 Then
        If TypeOf Source Is ListView And Dragging <> EndDrag Then ListView1.DragIcon = ImageList3.ListImages(1).Picture
        Err.Clear
    Else
        ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
    End If
    
End Sub

' This routine is called when the user expands a project in the
' TreeView control.

Private Sub TreeView1_Expand(ByVal Node As Node)

    Dim WorkingFolder As String
    Dim Response As Long
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Open the current Project glyph
    Node.Image = "Open"
    TreeView1.Nodes(Node.Index).Selected = True
    
    ' Clear the File List
    ListView1.ListItems.Clear

    ' Check to see if we are expanding the Root Project and
    ' instantiate the VSSItem as appropriate
    If TreeView1.SelectedItem.Key = "$" Then
        lblContents.Caption = "Contents of: $/"
    Else
        lblContents.Caption = "Contents of: " + TreeView1.SelectedItem.Key
    End If
    Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
    
    ' Determine and display working folder on main form
    Call DisplayWorkingFolder(objVSSProject)
    
    ' Call routine to populate the TreeView and ListView controls
    Call PopulateMain(objVSSProject)
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:

        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
            Err.Clear
        End If
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
End Sub
    
' This routine is passed a project item as a parameter. It checks for existing
' sub projects in the passed project and is used for populating the TreeView
' Control. If the passed project contains sub projects, it must be added to
' the control allowing the project to be "expanded"

Public Sub PopulateSubProjects(objVSSProject As VSSItem)

    Dim Nodex As Node
    Dim objVSSObject As VSSItem
    
    ' Iterate through each item of the project (false means ignore deleted)
    For Each objVSSObject In objVSSProject.Items(False)
        
        ' If a sub project (type = 0) is found then add it as a child to the
        ' current project
        If objVSSObject.Type = 0 Then
            Set Nodex = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key + "/" + objVSSProject.Name, tvwChild, TreeView1.SelectedItem.Key + "/" + objVSSProject.Name + "/" + objVSSObject.Name, objVSSObject.Name, "Closed")
        End If
    Next
    
End Sub

' This routine is called when the user elects to check out a file. It
' checks to see how many files are to be checked out and
' initializes the CheckOut dialog as appropriate.
    
Public Sub CheckOut()

    Dim Count As Integer
    Dim FileCount As Integer
    Dim Response As Long
    Dim FoundSelectedFile As Boolean
    Dim WorkingDirectory As String
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    FoundSelectedFile = False
    FileCount = 0

    ' Checking Out a file(s)
    If Selected = "Listview1" Then
        
        ' Instantiate the selected file
        Set objVSSProject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
            
        ' Iterate through each item in the ListView Control
        For Count = 1 To ListView1.ListItems.Count
            
            ' File is selected
            If ListView1.ListItems(Count).Selected = True Then
                FoundSelectedFile = True
                
                ' Tally total files selected
                FileCount = FileCount + 1
                
                ' Set appropriate caption for Check Out dialog based on
                ' number of files to be checked out
                If FileCount > 1 Then
                    frmCheckOut.Caption = "Check Out Multiple"
                    Exit For
                Else
                    frmCheckOut.Caption = "Check Out " + objVSSProject.Name
                End If
            End If
        Next
        
        ' If no selected files are found then check for a
        ' default selection
        If Not FoundSelectedFile Then
            If ListView1.SelectedItem.Key <> "" Then
                ListView1.ListItems(ListView1.SelectedItem.Key).Selected = True
                frmCheckOut.Caption = "Check Out " + objVSSProject.Name
                FoundSelectedFile = True
            End If
        End If
        
        ' If there are files to be checked out then show the Check Out Dialog
        If FoundSelectedFile Then
            frmCheckOut.Show 1
        Else
            Response = MsgBox("No file(s) selected", vbExclamation, AppTitle)
        End If
            
    ' Checking out a project
    Else
    
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Set the CheckOut dialog caption
        frmCheckOut.Caption = "Check Out " + objVSSProject.Spec
        frmCheckOut.chkRecursive.Visible = True
        frmCheckOut.Show 1
        
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
                
        ' If error is due to not having a working folder set or a non unique key
        ' then don't display this message
        If Err <> 364 And Err <> 35602 Then
            Response = MsgBox("Unable to CheckOut file(s)." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' This routine is called when the user elects to check in a file.
' It checks to see how many files are to be checked in and
' initializes the CheckIn dialog as appropriate.
    
Public Sub CheckIn()

    Dim Count As Integer
    Dim FileCount As Integer
    Dim Response As Long
    Dim FoundSelectedFile As Boolean
    Dim CheckOutDirectory As String
    Dim objVSSFile As VSSItem
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error Resume Next
    
    ' Initialize variables
    Count = 0
    FoundSelectedFile = False
    
    ' CheckingIn a file(s)
    If Selected = "Listview1" Then
    
        ' Instantiate the selected file
        Set objVSSFile = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Iterate through each item in the ListView Control
        For FileCount = 1 To ListView1.ListItems.Count
            
            ' File is selected
            If ListView1.ListItems(FileCount).Selected = True Then
                FoundSelectedFile = True
               
                ' Tally total files selected
                Count = Count + 1
                
                ' Initialize Check In dialog based on number
                ' of files to be checked in
                If Count > 1 Then
                    frmCheckIn.Caption = "Check In Multiple"
                    frmCheckIn.txtCheckInFrom.Text = ""
                    frmCheckIn.cmdDiff.Enabled = False
                    Exit For
                Else
                    CheckOutDirectory = GetVSSPath(objVSSFile.LocalSpec)
                    frmCheckIn.Caption = "Check In " + objVSSFile.Name
                    frmCheckIn.txtCheckInFrom.Text = CheckOutDirectory
                End If
            End If
        Next
    
        ' If no selected files are found then check for a
        ' default selection
        If Not FoundSelectedFile Then
            If ListView1.SelectedItem.Key <> "" Then
                ListView1.ListItems(ListView1.SelectedItem.Key).Selected = True
                frmCheckIn.Caption = "Check In " + objVSSFile.Name
                FoundSelectedFile = True
                
                ' Check for a comment that may have been applied when the file
                ' was Checked Out and find the directory it was checked out to
                For Each objVSSCheckout In objVSSFile.Checkouts
                    CheckOutDirectory = objVSSCheckout.LocalSpec
                    frmCheckIn.txtComment.Text = objVSSCheckout.Comment
                Next
            End If
        End If
        
        ' If there are files to be checked in then show the Check In Dialog
        If FoundSelectedFile Then
            frmCheckIn.Show 1
        Else
            Response = MsgBox("No file(s) selected", vbExclamation, AppTitle)
        End If
        
    ' Checking in a project
    Else
    
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Set the CheckOut dialog caption
        frmCheckIn.Caption = "Check In " + objVSSProject.Spec
        frmCheckIn.txtCheckInFrom.Text = objVSSProject.LocalSpec
        frmCheckIn.chkRecursive.Visible = True
        frmCheckIn.Show 1
    End If
    
    ' Set mousepointer
    MousePointer = vbNormal
        
End Sub

' This routine is called when the user elects to get a file.
' It checks to see how many files are to be gotten and
' initializes the Get dialog as appropriate.
    
Public Sub GetFile()

    Dim Count As Integer
    Dim FileCount As Integer
    Dim Response As Long
    Dim FoundSelectedFile As Boolean
    
    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    FoundSelectedFile = False
    
    ' Getting a file(s)
    If Selected = "Listview1" Then
    
        ' Instantiate the selected file
        Set objVSSProject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Set the make writeable checkbox to visible
        frmGet.chkMakeWriteable.Visible = True
        
        ' Iterate through each item in the ListView Control
        For Count = 1 To ListView1.ListItems.Count
            
            ' Tally total files selected
            If ListView1.ListItems(Count).Selected = True Then
                FoundSelectedFile = True
                FileCount = FileCount + 1
                
                ' Set appropriate caption for Get dialog based on number of
                ' files to be gotten
                If FileCount > 1 Then
                    frmGet.Caption = "Get Multiple"
                    Count = ListView1.ListItems.Count
                Else
                    frmGet.Caption = "Get " + objVSSProject.Name
                End If
            End If
        Next
        
        ' If no selected files are found then check for a
        ' default selection
        If Not FoundSelectedFile Then
            If ListView1.SelectedItem.Key <> "" Then
                ListView1.ListItems(ListView1.SelectedItem.Key).Selected = True
                frmGet.Caption = "Get " + objVSSProject.Name
                FoundSelectedFile = True
            End If
        End If
        
        ' If there are files to be gotten then show the Get Dialog
        If FoundSelectedFile Then
            frmGet.Show 1
        Else
            Response = MsgBox("No file(s) selected", vbExclamation, AppTitle)
        End If
        
    ' Getting a project
    Else
    
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Set the Get dialog caption
        frmGet.Caption = "Get " + objVSSProject.Spec
        If objVSSProject.LocalSpec <> "" Then
            frmGet.txtGetTarget.Text = objVSSProject.LocalSpec
        End If
        frmGet.chkRecursive.Visible = True
        frmGet.Show 1
    End If
        
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' No working folder set
        If Err = -2147166539 Then
            Resume Next
        Else
            Response = MsgBox("Unable to Get file(s)." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' This routine is called when the user elects to un-check out a file.
' It checks to see how many files are to be unchecked out and
' initializes the UndoCheckOut dialog as appropriate.
    
Public Sub UndoCheckOut()

    Dim FileCount As Integer
    Dim Count As Integer
    Dim Response As Long
    Dim FoundSelectedFile As Boolean
    
    ' Set MousePointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    FoundSelectedFile = False
    Count = 0
    
    ' Set On Error routine
    On Error Resume Next
    
    ' UnCheckout a file(s)
    If Selected = "Listview1" Then
    
        ' Instantiate the selected file
        Set objVSSProject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
    
        ' Iterate through each item in the ListView Control
        For FileCount = 1 To ListView1.ListItems.Count
            
            ' File is selected
            If ListView1.ListItems(FileCount).Selected = True Then
                FoundSelectedFile = True
               
               ' Tally total files selected
                Count = Count + 1
                
                ' Initialize the Undo CheckOut dialog
                If Count > 1 Then
                    frmUndoCheckOut.Caption = "Undo Check Out Multiple"
                    Exit For
                Else
                    frmUndoCheckOut.Caption = "Undo Check Out " + objVSSProject.Name
                End If
            End If
        Next
        
        ' If no selected files are found then check for a
        ' default selection
        If Not FoundSelectedFile Then
            If ListView1.SelectedItem.Key <> "" Then
                ListView1.ListItems(ListView1.SelectedItem.Key).Selected = True
                frmUndoCheckOut.Caption = "Undo Check Out " + objVSSProject.Name
                FoundSelectedFile = True
            End If
        End If
        
        ' If there are files to be unchecked out then show the UndoCheckout Dialog
        If FoundSelectedFile Then
            frmUndoCheckOut.Show 1
        Else
            Response = MsgBox("No file(s) selected", vbExclamation, AppTitle)
        End If
        
    ' UncheckOut a project
    Else
    
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Set the CheckOut dialog caption
        frmUndoCheckOut.Caption = "Undo Check Out " + objVSSProject.Spec
        frmUndoCheckOut.chkRecursive.Visible = True
        frmUndoCheckOut.Show 1
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' This routine is called when the user elects to view the history
' of a file or project.
    
Public Sub ShowHistory()

    ' Set Mousepointer
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    MousePointer = vbHourglass
    
    ' Set flag indicating whether we are showing file or project history
    FileHistory = (Selected = "Listview1")

    If Not FileHistory Then
    
        ' Open the Options dialog
        frmHistoryOptions.Show 1
    Else
        
        ' Show History Dialog
        frmHistoryOptions.cmdOK_Click
    End If
    
    ' Set Mousepointer
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    MousePointer = vbNormal
    
End Sub

' This routine is called when the user elects to delete a file or project
' (either via the menu or the toolbar. It checks to see if the selected
' item is a project or a file and initializes the Delete dialog as appropriate.
    
Public Sub DeleteItem()

    Dim Response As Long
    Dim ItemName As String
    Dim FoundSelectedFile As Boolean
    Dim FileCount As Integer
    Dim Count As Integer
    Dim objVSSFile As VSSItem
    
    ' Set On Error Routine
    On Error GoTo ErrHandler

    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    FoundSelectedFile = False
    FileCount = 0

    ' Deleting a project
    If Selected = "Treeview1" Then
        
        ' Check if user has selected the root
        If TreeView1.SelectedItem.Key = "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
            ItemName = "$/"
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
            ItemName = objVSSProject.Name
        End If
        
        ' Set Delete dialog properties
        frmDelete.Caption = "Delete item " + ItemName
        frmDelete.lblItem.Caption = "Delete project '" + ItemName + "' and all of it's contents?"
        
        ' Show Delete dialog
        frmDelete.Show 1
    
    ' Deleting a file
    ElseIf Selected = "Listview1" Then
    
        ' Set VSS Item to selected file in ListView control
        Set objVSSFile = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
    
        ' Set Delete dialog properties
        frmDelete.Caption = "Delete item " + objVSSFile.Name
        frmDelete.lblItem.Caption = "Delete file '" + objVSSFile.Name + "'?"

        ' Iterate through each item in the ListView Control
        For Count = 1 To ListView1.ListItems.Count
        
            ' Item is selected
            If ListView1.ListItems(Count).Selected = True Then
                FoundSelectedFile = True
                
                ' Tally total files selected
                FileCount = FileCount + 1
                
                ' Set appropriate caption for Check Out dialog based on
                ' number of files to be checked out
                If FileCount > 1 Then
                    frmDelete.Caption = "Delete Multiple"
                    frmDelete.lblItem.Caption = "Delete all selected items?"
                    Exit For
                End If
                
                ' Set Delete dialog properties
                frmDelete.Caption = "Delete item " + objVSSFile.Name
                frmDelete.lblItem.Caption = "Delete file '" + objVSSFile.Name + "'?"
            End If
        Next Count
        
        ' Show Delete dialog
        frmDelete.Show 1
        
        ' Update the GUI
        RefreshFileList
        
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then Response = MsgBox("Unable to delete item." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
        
        ' Set Mousepointer
        MousePointer = vbNormal
    
    End If
    
End Sub

' This routine is called when the user elects to label a file or
' project It checks to see if the selected item is a file or project
' and initalizes the Label dialog as appropriate.

Public Sub LabelItem()

    Dim ItemName As String
    
    ' Set Mousepointer
    MousePointer = vbHourglass

    ' Labeling a File
    If Selected = "Listview1" Then
    
        ' Set VSSItem to the selected file in the ListView control
        Set objVSSProject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Initialize the Label dialog
        frmLabel.Caption = "Label " + objVSSProject.Name
        frmLabel.lblItem.Caption = "Label File: " + objVSSProject.Name
        
    ' Labeling a Project
    ElseIf Selected = "Treeview1" Then
    
        ' Check if user is labeling the root ($/) and set the current
        ' VSS Item as appropriate
        If TreeView1.SelectedItem.Key = "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
            ItemName = "$/"
        Else
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
            ItemName = objVSSProject.Spec
        End If
        
        ' Initialize the Label dialog
        frmLabel.Caption = "Label " + ItemName
        frmLabel.lblItem.Caption = "Label Project: " + ItemName

    End If
    
    ' Show the dialog
    frmLabel.Show 1
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' This routine is called when the user elects to add files to a
' selected project

Public Sub AddFile()

    ' Check to see if user is adding files to the root project.
    ' Set the current VSS Item as appropriate, initialize the
    ' Add File dialog and display it.
    If TreeView1.SelectedItem.Key = "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        frmAddFile.Caption = "Add Items to $/"
        frmAddFile.Show 1
    ElseIf TreeView1.SelectedItem.Key <> "$" Then
        frmAddFile.Caption = "Add Items to " + objVSSProject.Name
        frmAddFile.Show 1
    End If
    
End Sub

' This routine is called when the user elects to add
' a new project to the database

Public Sub CreateProject()
    
    Dim Response As Long

    ' Check to see if user is adding a project to the root
    ' project. Set the current VSS Item as appropriate
    ' initialize the Add Project dialog and show it
    If TreeView1.SelectedItem.Key = "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        frmCreateProject.Caption = "Create Project in " + "$/"
        frmCreateProject.Show 1
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        frmCreateProject.Caption = "Create Project in " + TreeView1.SelectedItem.Key
        frmCreateProject.Show 1
    End If
    
End Sub

' This routine is called when the user elects to Rename a selected project
' or file from the menu. It determines if the selected item is a file or a
' project and then initializes the Rename dialog as appropriate
    
Public Sub Rename()

    Dim Response As Long
    Dim ItemName As String
    
    ' Renaming a file
    If Selected = "Listview1" Then
    
        ' Set VSS Item to the selected file
        Set objVSSProject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Check to see if file is checked out
        If objVSSProject.IsCheckedOut <> VSSFILE_NOTCHECKEDOUT Then
            Response = MsgBox("Sorry, you cannot rename a file that is checkout.", vbInformation, AppTitle)
            Exit Sub
        End If
        
        ' Initialize the Rename dialog
        frmRename.Caption = "Rename File: " + frmMain.ListView1.SelectedItem.Key
    
    ' Renaming a project
    ElseIf Selected = "Treeview1" Then
        
        ' Set VSS Item to the selected Project
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
            ItemName = objVSSProject.Name
        Else
            Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
            ItemName = "$\"
        End If
        
        ' Initialize the Rename dialog
        frmRename.Caption = "Rename Project: " + ItemName
    End If
    
    ' Show the Rename dialog
    frmRename.Show 1
    
End Sub
    
' This routine is called when the user elects to set the Working directory
' for the selected file or project
    
Public Sub SetWorkingFolder()
    
    ' Set VSS Item to the selected Project
    If TreeView1.SelectedItem.Key = "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem("$/", False)
        frmWorkingFolder.Caption = "Set Working Folder for $/"
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        frmWorkingFolder.Caption = "Set Working Folder for " + TreeView1.SelectedItem.Key
    End If
    
    ' Show the Set Working Dir dialog
    frmWorkingFolder.Show 1
    
End Sub

' This routine refreshes the TreeView and ListView controls to
' reflect and changes made to the database by other users since
' the control was last populated.
    
Public Sub RefreshFileList()

    Dim Nodex As Node
    Dim ItemCount As Integer
    Dim WorkingFolder As String
    Dim objVSSProject As VSSItem
    
    ' Set Mouespointer
    MousePointer = vbHourglass
    frmMain.StatusBar1.Panels(1).Text = "Working..."
    
    ' Set On Error routine
    On Error Resume Next

    ' Clear the files from the ListView Control
    frmMain.ListView1.ListItems.Clear
    
    ' Set the selected Project Item
    If TreeView1.SelectedItem.Key <> "$" Then
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        lblContents.Caption = "Contents of: " + TreeView1.SelectedItem.Key
    Else
        Set objVSSProject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        lblContents.Caption = "Contents of: $/"
    End If
    
    ' Find working folder for Project and display it
    Call DisplayWorkingFolder(objVSSProject)
    
    ' Call routine to populate TreView and ListView Controls
    Call PopulateMain(objVSSProject)
    
    ' Set Mouespointer
    MousePointer = vbNormal
    frmMain.StatusBar1.Panels(1).Text = "Ready"
    
End Sub
    
' This routine is called when the user has selected to (either from the
' menu or the toolbar) to view the properties of a file or project. This
' routine determines if the user has selected a file or a project and
' initializes/populates the Property dialog as appropriate

Public Sub Properties()

    Dim Response As Long

    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Properties for a Project
    If Selected = "Treeview1" Then
        
        ' Check to see if user wants properties for the root project
        ' and set the VSS Item as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objVSSObject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
            
            ' Initialize Property dialog
            frmProperties.lblName.Caption = "Name: " + objVSSObject.Name
            frmProperties.Caption = frmMain.TreeView1.SelectedItem.Key
        Else
            Set objVSSObject = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
            
            ' Initialize Property dialog
            frmProperties.lblName.Caption = "Name: $/"
            frmProperties.Caption = "$/"
        End If
        
    ' Properties for a File
    ElseIf Selected = "Listview1" Then
    
        ' Set VSS Item to the currently selected file
        Set objVSSObject = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Initialize Property dialog
        frmProperties.Caption = frmMain.ListView1.SelectedItem.Key
        frmProperties.lblName.Caption = "Name: " + objVSSObject.Name
    End If
    
    ' Call routine to populate the Properties dialog
    Call PopulateProperties
        
    ' Set Mousepointer
    MousePointer = vbNormal
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then Response = MsgBox("Unable to display properties." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
        
    ' Set Mousepointer
    MousePointer = vbNormal
        
End Sub

' This function populate the Properties form base on the type
' of VSSItem selected

Public Sub PopulateProperties()
    
    Dim NodeNumber As Integer
    Dim objVSSProjectSubItem As VSSItem
    Dim objVSSLinkItem As VSSItem
    Dim ProjectCount As Integer
    Dim FileCount As Integer
    Dim CheckOutCount As Integer
    Dim objVSSCheckout As VSSCheckout
    Dim objUserListItem As ListItem
    
    ' Initialize variables
    CheckOutCount = 0
    FileCount = 0
    ProjectCount = 0
    
    ' Initialize Property dialog Project
    If objVSSObject.Type = VSSITEM_PROJECT Then
        frmProperties.cmbFileType.Visible = False
        frmProperties.lblType.Caption = "Type: Project"
    End If
    
    ' Iterate through version collection of current VSS Item to obtain
    ' current version and date. Since this will be the first version, get
    ' that data and exit For loop
    For Each objVSSVersion In objVSSObject.Versions
        frmProperties.txtComment.Text = objVSSVersion.Comment
        frmProperties.lblVersion.Caption = "Version: " + Str(objVSSVersion.VersionNumber)
        frmProperties.lblVersionDate.Caption = "Date: " + Str(objVSSVersion.Date)
        Exit For
    Next
    
    ' Iterate through version collection of current VSS Item to get latest
    ' Label (if any). Once that is obtained, set variables and exit For loop
    For Each objVSSVersion In objVSSObject.Versions
        If objVSSVersion.Label <> "" Then
            frmProperties.lblLastLabel.Caption = "Last label: " + objVSSVersion.Label
            frmProperties.lblLabelVersion.Caption = "Version: " + Str(objVSSVersion.VersionNumber)
            frmProperties.lblLabelDate.Caption = "Date: " + Str(objVSSVersion.Date)
            Exit For
        End If
    Next
    
    ' Show Properties for a file
    If objVSSObject.Type = VSSITEM_FILE Then
        
        ' If file is Checked out setup Checked out data
        If objVSSObject.IsCheckedOut Then
        
            ' Check for multiple checkouts
            For Each objVSSCheckout In objVSSObject.Checkouts
                CheckOutCount = CheckOutCount + 1
                If CheckOutCount > 1 Then Exit For
            Next
            
            ' File is checked out by one user
            If CheckOutCount = 1 Then
                frmProperties.lblFileNotCheckedOut.Caption = "Checked out"
                frmProperties.lblBy.Visible = True
                frmProperties.lblCheckOutDate.Visible = True
                frmProperties.lblCheckOutVersion.Visible = True
                frmProperties.lblCheckOutSystem.Visible = True
                frmProperties.lblCheckOutFolder.Visible = True
                frmProperties.lblCheckOutProject.Visible = True
                frmProperties.lblFileNotCheckedOut.Visible = True
                frmProperties.lblComment.Visible = True
                frmProperties.txtCheckOutComment.Visible = True
                For Each objVSSCheckout In objVSSObject.Checkouts
                    frmProperties.lblBy.Caption = "By: " + objVSSCheckout.UserName
                    frmProperties.lblCheckOutDate.Caption = "Date: " + Str(objVSSCheckout.Date)
                    frmProperties.lblCheckOutVersion.Caption = "Version: " + Str(objVSSCheckout.VersionNumber)
                    frmProperties.lblCheckOutSystem.Caption = "Computer: " + objVSSCheckout.Machine
                    frmProperties.lblCheckOutFolder.Caption = "Folder: " + objVSSCheckout.LocalSpec
                    frmProperties.lblCheckOutProject.Caption = "Project: " + objVSSCheckout.Project
                    frmProperties.txtCheckOutComment.Text = objVSSCheckout.Comment
                Next
            Else
                frmProperties.lblFileNotCheckedOut.Caption = "Check Outs:"
                frmProperties.CheckedOutList.Visible = True
                
                ' Populate multiple checkout list
                For Each objVSSCheckout In objVSSObject.Checkouts
                    
                    Set objUserListItem = frmProperties.CheckedOutList.ListItems.Add(, objVSSCheckout.UserName, objVSSCheckout.UserName)
                    objUserListItem.SubItems(1) = objVSSCheckout.LocalSpec
                    
                Next
                frmProperties.cmdRecover.Visible = True
                frmProperties.cmdRecover.Caption = "Details"
                frmProperties.cmdRecover.Enabled = True
                Set frmProperties.CheckedOutList.SelectedItem = frmProperties.CheckedOutList.ListItems(1)
            End If
        
        ' File is not checked out
        Else
            frmProperties.lblFileNotCheckedOut.Caption = "Not checked out"
        End If
        
        ' Populate Links information
        frmProperties.lblLinksFor.Caption = "Projects sharing " + objVSSObject.Name
        For Each objVSSLinkItem In objVSSObject.Links
            frmProperties.lstLinks.AddItem (objVSSLinkItem.Parent.Spec)
        Next
    
    ' Show properties for a project
    Else
        
        ' Get the project information
        If TreeView1.Nodes(frmMain.TreeView1.SelectedItem.Index).Children > 0 Then
            
            ' Get first child's text, and set NodeNumber to its index value.
            NodeNumber = frmMain.TreeView1.SelectedItem.Child.Index
    
            ' While NodeNumber is not the index of the child node's
            ' last sibling, check the item type
            While NodeNumber <= TreeView1.SelectedItem.Child.LastSibling.Index
    
                Set objVSSProjectSubItem = objVSSDatabase.VSSItem(TreeView1.Nodes(NodeNumber).Key, False)
                If objVSSProjectSubItem.Type = VSSITEM_PROJECT Then ProjectCount = ProjectCount + 1
                If NodeNumber < TreeView1.SelectedItem.Child.LastSibling.Index Then
                    NodeNumber = TreeView1.Nodes(NodeNumber).Next.Index
                Else
                    NodeNumber = NodeNumber + 1
                End If
            Wend
            frmProperties.lblProjects.Caption = ProjectCount & " Projects"
        Else
            frmProperties.lblProjects.Caption = "0 Projects"
        End If
        frmProperties.lblFiles.Caption = ListView1.ListItems.Count & " Files"
        frmProperties.lblContains.Visible = True
        frmProperties.lblProjects.Visible = True
        frmProperties.lblFiles.Visible = True
        frmProperties.SSTab1.TabEnabled(2) = False
        frmProperties.SSTab1.TabCaption(1) = "Deleted Items"
        Call CheckForDeletedItems(objVSSObject)
    End If
    
    ' Show Properties dialog
    frmProperties.Show 1
    
End Sub

' This routine is called when the user or another routine has asked
' to view a file's contents. It finds the Windows TEMP dir and Gets
' a copy to that dir. It then shells the file into NotePad
    
Public Sub ViewFile(FileKey As String)

    Dim lpBuffer As String * 100
    Dim TempPath As String
    Dim Response As Long
    Dim Count As Integer
    Dim RetVal As Long
    Dim objVSSFile As VSSItem
    
    ' Set Mousepointer
    MousePointer = vbHourglass

    ' Initialize variable
    Count = 1
    lpBuffer = String(100, Chr(0))
    Response = vbYes
    
    ' Set on error routine
    On Error GoTo ErrHandler
    
    ' Set current VSS Item to selected file in ListView control
    Set objVSSFile = objVSSDatabase.VSSItem(FileKey, False)
    If Err = 0 Then
        
        ' Warn user if file is binary
        If objVSSFile.Binary = True Then Response = MsgBox("The file " + objVSSFile.Name + " is binary. Continue?", vbYesNo + vbQuestion, AppTitle)
        
        ' User wants to view file
        If Response = vbYes Then
        
            ' Get the file to the Windows Temp dir and display it
            RetVal = GetTempPath(100, lpBuffer)
            TempPath = Mid(lpBuffer, 1, RetVal)
            If Mid(TempPath, Len(TempPath), 1) = "\" Then
                TempPath = TempPath + objVSSFile.Name
            Else
                TempPath = TempPath + "\" + objVSSFile.Name
            End If
            objVSSFile.Get Local:=TempPath, iFlags:=0
            RetVal = Shell("NotePad " + TempPath, vbNormalFocus)
        End If
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        Response = MsgBox("Unable to view file!" + vbCrLf + Err.Description, vbExclamation, AppTitle)
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub
    
Sub ViewFileOCX()
    
    Dim objVSSFile As VSSItem
    Dim GetDirectory As String
    Dim lpFileData As String * 100
    Dim TempPath As String
    Dim Count As Integer
    Dim Response As Long
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    lpFileData = String(100, Chr(0))
    TempPath = GetTempPath(100, lpFileData)
    GetDirectory = Mid(lpFileData, 1, TempPath)
    
    ' Instantiate the selected item in the History window
    Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
    
    ' Set Temp Dir Path for getting file
    If Mid(GetDirectory, Len(GetDirectory), 1) = "\" Then
        GetDirectory = GetDirectory + objVSSFile.Name
    Else
        GetDirectory = GetDirectory + "\" + objVSSFile.Name
    End If
    
    ' Get the file
    objVSSFile.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
    
    ' View the file
    Viewer1.ViewFile (GetDirectory)
    
    ' Check for errors
    If Err <> 0 Then
    
ErrHandler:

        Response = MsgBox("Cannot view file." + Err.Description, vbExclamation, AppTitle)
    End If

End Sub

' This routine is called when the user has asked to edit a file
' from the Menu. It checks to see if the file is already checked
' out and, if it is, shells that writeable copy into NotePad. Otherwise
' it Checks the file out to the Working Dir and then shells the file
' into NotePad.
    
Public Sub EditFile()

    Dim RetVal As Long
    Dim Response As Long
    Dim objVSSFile As VSSItem
    Dim WorkingFolder As String

    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    Response = vbYes
   
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set current VSS Item to selected file in ListView control
    Set objVSSFile = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
    If Err = 0 Then
        
        ' Warn user if file is binary
        If objVSSFile.Binary = True Then Response = MsgBox("The file " + objVSSFile.Name + " is binary. Continue?", vbYesNo + vbQuestion)

        ' User wants to edit file
        If Response = vbYes Then
        
            ' File is already checked out by user
            If objVSSFile.IsCheckedOut = VSSFILE_CHECKEDOUT_ME Then
                For Each objVSSCheckout In objVSSFile.Checkouts
                    If objVSSCheckout.UserName = UserName Then
                        WorkingFolder = objVSSCheckout.LocalSpec
                    End If
                Next
                
                ' Display writeable copy in NotePad
                If Right(WorkingFolder, 1) <> "\" Then WorkingFolder = WorkingFolder + "\"
                RetVal = Shell("NotePad " + WorkingFolder + objVSSFile.Name, vbNormalFocus)
            
            ' Check the file out to the working directory
            Else
                FileDate = Date + Time
                
                ' Check for Working Directory
                On Error Resume Next
                WorkingFolder = objVSSFile.Parent.LocalSpec
                If WorkingFolder = "" Then
                    Response = MsgBox("You must set a working folder for this file before it can be edited." + vbCrLf + "Would you like to set one now?", vbExclamation + vbYesNo, AppTitle)
                    If Response = vbYes Then Call SetWorkingFolder
                Else
                
                    ' CheckOut the file and edit in NotePad
                    On Error GoTo ErrHandler
                    If Right(WorkingFolder, 1) <> "\" Then WorkingFolder = WorkingFolder + "\"
                    objVSSFile.CheckOut Comment:="", Local:=WorkingFolder + objVSSFile.Name, iFlags:=EOLFlag + VSSFLAG_CHKEXCLUSIVENO
                    frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " checked out" + vbCrLf
                    Call ShowResults
                    
                    ' Display writeable copy in NotePad
                    RetVal = Shell("NotePad " + WorkingFolder + objVSSFile.Name, vbNormalFocus)
                End If
            End If
        End If
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' No working folder set
        If Err <> -2147166539 Then
            Response = MsgBox("Unable to edit file!" + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
    End If
    
    '  Update GUI and Set Mousepointer
    RefreshFileList
    MousePointer = vbNormal
    
End Sub

' This routine adds an item to the ListView control representing a file that is
' checked out

Public Sub PopulateFileItem(CheckedOut As Boolean, FileIcon As String, Optional CheckedOutBy As Variant, Optional CheckOutDirectory As Variant)

    Dim objVSSFile As ListItem
    Dim strDate As String
    
    ' Add the item to the ListView
    If CheckedOut Then
        Set objVSSFile = ListView1.ListItems.Add(, TreeView1.SelectedItem.Key + "/" + objVSSObject.Name, objVSSObject.Name, , FileIcon)
        objVSSFile.SubItems(1) = CheckedOutBy
        objVSSFile.SubItems(3) = CheckOutDirectory
    Else
        Set objVSSFile = ListView1.ListItems.Add(, TreeView1.SelectedItem.Key + "/" + objVSSObject.Name, objVSSObject.Name, , FileIcon)
    End If
    
    ' Set the date
    objVSSFile.SubItems(2) = Format(FileDate, "mm/dd/yy hh:mm ampm")

End Sub

' This routine adds an item to the ListView control representing a
' file.

Public Sub AddFileItem(FileToAdd As String, VSSPath As Boolean, FileIcon As String)
    
    Dim FileItem As ListItem
    Dim FileName As String
    
    ' Get the FileName of the File to Add
    If VSSPath Then
        FileName = GetVSSFileName(FileToAdd)
    Else
        FileName = GetFileName(FileToAdd)
    End If
    
    ' Add item to FileList
    Set FileItem = frmMain.ListView1.ListItems.Add(, FileToAdd, FileName, , FileIcon)
    frmMain.ListView1.Sorted = True
    FileItem.SubItems(1) = ""
    FileItem.SubItems(2) = FileDate
    FileItem.SubItems(3) = ""
    
End Sub
    
' Project list is selected

Private Sub TreeView1_GotFocus()

    Selected = "Treeview1"
    mnuBranch.Enabled = False
    Toolbar1.Buttons.Item("Branch").Enabled = False
    
End Sub

' This routine sorts the ListView control by Filename

Public Sub SortByName()
    
    mnuName.Checked = True
    mnuUser.Checked = False
    mnuDate.Checked = False
    mnuCheckOutFolder.Checked = False
    ListView1.SortKey = 0
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True
    StatusBar1.Panels(4).Text = "Sort: Name"
    
End Sub

' This routine sorts the ListView control by Username
    
Public Sub SortByUser()

    mnuName.Checked = False
    mnuUser.Checked = True
    mnuDate.Checked = False
    mnuCheckOutFolder.Checked = False
    ListView1.SortKey = 1
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True
    StatusBar1.Panels(4).Text = "Sort: User"
    
End Sub

' This routine sorts the ListView control by Date

Public Sub SortByDate()
    
    mnuName.Checked = False
    mnuUser.Checked = False
    mnuDate.Checked = True
    mnuCheckOutFolder.Checked = False
    ListView1.SortKey = 2
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True
    StatusBar1.Panels(4).Text = "Sort: Date"
    
End Sub

' This routine sorts the ListView control by the CheckOut Folder
    
Public Sub SortByFolder()

    mnuName.Checked = False
    mnuUser.Checked = False
    mnuDate.Checked = False
    mnuCheckOutFolder.Checked = True
    ListView1.SortKey = 3
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True
    StatusBar1.Panels(4).Text = "Sort: Folder"
    
End Sub

' This routine inverts the sort order for the ListView control

Public Sub InvertSortOrder()
    
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    
End Sub
    
' This routine inverts the selected items in the ListView control

Public Sub InvertSelection()

    Dim Count As Integer
    
    For Count = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Count).Selected = True Then
            ListView1.ListItems(Count).Selected = False
        Else
            ListView1.ListItems(Count).Selected = True
        End If
    Next
    ListView1.SetFocus
    
End Sub

' This procedure allows the user to share files between projects

Public Sub Share()

    ' Check if user is Sharing to the root dir and set the VSS Item appropriately
    If TreeView1.SelectedItem.Key = "$" Then
        frmShare.Caption = "Share with $/"
    Else
        frmShare.Caption = "Share with " + TreeView1.SelectedItem.Key
    End If
    
    ' Show the Share dialog
    frmShare.Show 1
    
    ' Update the GUI
    RefreshFileList
    
End Sub

Public Sub ShowDifferences()
    
    Dim FileIsDifferent As Boolean
    Dim Response As Long
    Dim CheckOutFolder As String
    Dim objVSSFile As VSSItem
    Dim GetDirectory As String
    Dim lpFileData As String * 100
    Dim TempPath As String
    Dim Count As Integer
    
    ' Set On Error Routine
    On Error GoTo ErrHandler
    
    ' Initialize variables
    lpFileData = String(100, Chr(0))
    TempPath = GetTempPath(100, lpFileData)
    GetDirectory = Mid(lpFileData, 1, TempPath)
    
    ' Set the Mousepointer
    MousePointer = vbHourglass
    
    ' Diffing a file
    If Selected = "Listview1" Then
    
        ' Instantiate the selected file
        Set objVSSFile = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        CheckOutFolder = GetLocalCheckOutFolder(objVSSFile)
        
        If Not DiffMethod Then
        
            ' Compare local file to VSS file
            FileIsDifferent = objVSSFile.IsDifferent(Local:=CheckOutFolder)
            
            Call ShowDiffInfo(FileIsDifferent, objVSSFile.Binary)
        Else
           
    
            ' Instantiate the selected item in the History window
            Set objVSSFile = objVSSDatabase.VSSItem(frmMain.ListView1.SelectedItem.Key, False)
    
            ' Set Temp Dir Path for getting file
            If Mid(GetDirectory, Len(GetDirectory), 1) = "\" Then
                GetDirectory = GetDirectory + "SourceSafe version"
            Else
                GetDirectory = GetDirectory + "\SourceSafe version"
            End If
    
            ' Get the file
            objVSSFile.Get Local:=GetDirectory, iFlags:=VSSFLAG_REPREPLACE
            Diff1.DiffTwoFiles GetDirectory, CheckOutFolder, frmMain.hWnd
        End If
    Else
       
        ' The IsDifferent method currently works on files only so
        Response = MsgBox("Sorry, this command only works on files.", vbExclamation, AppTitle)
    
    End If
    
    ' Check For Errors
    If Err <> 0 Then
ErrHandler:
    
        Response = MsgBox("Unable to show differences." + vbCrLf + Err.Description, vbExclamation, AppTitle)
        Err.Clear
    End If
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

' Opens the About Box

Sub About()

    frmAbout.Show 1

End Sub

' This routine populates the ListView and TreeView controls as appropriate

Public Sub PopulateMain(objVSSProject As Object)

    Dim Nodex As Node
    Dim WorkingFolder As String
    Dim objVSSLinkItem As VSSItem
    Dim FileIcon As String
    Dim LinkCount As Integer
    Dim OtherUser As String
    Dim Response As Long
    Dim objVSSCheckoutOther As VSSCheckout
    Dim UserNameProp As String
    Dim FileShared As Boolean
    
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set Item Count
    ItemCount = 0
    LinkCount = 0
    
    ' Initialize variables
    FileShared = False
    
    ' Find working folder for Project and display it
    Call DisplayWorkingFolder(objVSSProject)
    lblContents.Caption = "Contents of: " + objVSSProject.Spec
    
    ' Set the Current Project
    objVSSDatabase.CurrentProject = objVSSProject.Spec
    
    ' Iterate through all items in current project (false means ignore deleted items)
    For Each objVSSObject In objVSSProject.Items(False)
            
        ' Reset flag
        FileShared = False
            
        ' Check to see what type of object we have
        Select Case objVSSObject.Type
                    
            ' Current item is a project
            Case 0
            
                ' Although there is certainly a better way to do this,
                ' I use the ON ERROR to avoid recreating a node in the
                ' TreeView Control
                On Error Resume Next
                
                ' Add project to Treeview control
                Set Nodex = TreeView1.Nodes.Add("$", tvwChild, TreeView1.SelectedItem.Key + "/" + objVSSObject.Name, objVSSObject.Name, "Closed")
                
                ' Call procedure to check for existing sub projects of this
                ' project (this poplates the control with + signs as needed)
                Call PopulateSubProjects(objVSSObject)
                
                If Err = 0 Then Err.Clear
            
            ' Current Object is a file
            Case 1
                
                ' Iterate through the version collection to find file-date of current
                ' version. Since this will be the first version, get data on first pass
                ' and then exit for loop
                For Each objVSSVersion In objVSSObject.Versions
                    If objVSSVersion.VersionNumber = objVSSObject.VersionNumber Then
                        FileDate = objVSSVersion.Date
                        Exit For
                    End If
                Next
                
                ' Check File CheckoutStatus and populate as appropriate
                Select Case objVSSObject.IsCheckedOut
                    
                    ' File is checked out by current user
                    Case VSSFILE_CHECKEDOUT_ME
                    
                        ' Set Current UserName to be used if file is
                        ' Checked out by multiple users
                        UserNameProp = UserName
                    
                        ' Get the CheckOut directory
                        For Each objVSSCheckout In objVSSObject.Checkouts
                            OtherUser = objVSSCheckout.UserName
                            WorkingFolder = objVSSCheckout.LocalSpec
                            Exit For
                        Next
                        
                        ' Set File Icon
                        FileIcon = "Checked"
                        
                        ' Check to see if file is shared
                        For Each objVSSLinkItem In objVSSObject.Links
                            LinkCount = LinkCount + 1
                            If LinkCount = 2 Then
                                FileIcon = "CheckedOutShared"
                                FileShared = True
                                Exit For
                            End If
                        Next
                        
                        ' Check to see if file is checked out muliple
                        For Each objVSSCheckoutOther In objVSSObject.Checkouts
                            If objVSSCheckoutOther.UserName <> UserName Then
                                If FileShared Then
                                    FileIcon = "CheckoutSharedMulti"
                                Else
                                    FileIcon = "CheckedOutMulti"
                                End If
                                UserNameProp = UserName + "..."
                                Exit For
                            End If
                        Next
                        Call PopulateFileItem(True, FileIcon, UserNameProp, WorkingFolder)
                        
                    ' File is checked out by other user
                    Case VSSFILE_CHECKEDOUT
                        
                        ' Find username and CheckOut directory. Since this will be the
                        ' first version, get data on first pass and then exit for loop
                        For Each objVSSCheckout In objVSSObject.Checkouts
                            OtherUser = objVSSCheckout.UserName
                            WorkingFolder = objVSSCheckout.LocalSpec
                            Exit For
                        Next
                        FileIcon = "Checked"
                        
                        ' Check to see if file is shared
                        For Each objVSSLinkItem In objVSSObject.Links
                            LinkCount = LinkCount + 1
                            If LinkCount = 2 Then
                                FileIcon = "CheckedOutShared"
                                FileShared = True
                                Exit For
                            End If
                        Next
                        
                        ' Check for multiple checkouts
                        For Each objVSSCheckoutOther In objVSSObject.Checkouts
                            If objVSSCheckoutOther.UserName <> OtherUser Then
                                FileIcon = "CheckedOutMulti"
                                OtherUser = OtherUser + "..."
                                If FileShared Then
                                    FileIcon = "CheckoutSharedMulti"
                                Else
                                    FileIcon = "CheckedOutMulti"
                                End If
                                Exit For
                            End If
                        Next
                        
                        ' Populate listview with current item
                        Call PopulateFileItem(True, FileIcon, OtherUser, WorkingFolder)
                    
                    ' File is not checked out
                    Case Else
                        
                        ' Set File Icon
                        FileIcon = "Leaf"
                        
                        ' Check to see if file is shared
                        For Each objVSSLinkItem In objVSSObject.Links
                            LinkCount = LinkCount + 1
                            If LinkCount = 2 Then
                                FileIcon = "Shared"
                                Exit For
                            End If
                        Next
                        
                        ' Populate listview with current item
                        Call PopulateFileItem(False, FileIcon)
                        
                End Select
                
                ' Tally number of items in project for display purposes
                ItemCount = ItemCount + 1
            
            ' Unknown object type
            Case Else
                MsgBox ("Unknown object type encountered during Node population!")
        End Select
        
        ' Reset Link Count
        LinkCount = 0
    Next
    
    ' Update Status Bar with item count of current project
    Select Case Str(ItemCount)
        Case 1
            StatusBar1.Panels(5).Text = (Str(ItemCount) + " item")
        Case Else
            StatusBar1.Panels(5).Text = (Str(ItemCount) + " items")
    End Select
    
    If Err <> 0 Then
    
ErrHandler:
        
        ' We may get an error regarding a non unique key here which may be ignored
        If Err <> 35602 Then
            Response = MsgBox(Err.Description, vbExclamation, AppTitle)
        Else
            Resume
        End If
        Err.Clear
    End If
    
End Sub

' Iterates through the current project item looking for deleted items

Public Sub CheckForDeletedItems(objVSSProjectItem As Object)

    Dim objVSSSubItem As VSSItem
    Dim DeletedItemFound As Boolean
    
    ' Initialize variables
    DeletedItemFound = False
    
    ' Iterate through the project items collection and look for deleted items
    For Each objVSSSubItem In objVSSProjectItem.Items(IncludeDeleted:=True)
        
        ' If a deleted item is found add it to the List control of the Properties dialog
        If objVSSSubItem.Deleted = True Then
            DeletedItemFound = True
            frmProperties.lstDeleted.AddItem objVSSSubItem.Spec
        End If
    Next
    If DeletedItemFound Then
        frmProperties.lblFileNotCheckedOut.Caption = "Items:"
        frmProperties.lstDeleted.Visible = True
        frmProperties.cmdRecover.Visible = True
        frmProperties.cmdPurge.Visible = True
    
    ' No deleted items found
    Else
        frmProperties.lblFileNotCheckedOut.Caption = "There are no deleted items for this project"
        frmProperties.lstDeleted.Visible = False
        frmProperties.cmdRecover.Visible = False
        frmProperties.cmdPurge.Visible = False
    End If
    
End Sub

' This routine branches a shared file

Public Sub Branch(ItemToBranch As String)
    
    Dim FileCount As Integer
    Dim Response As Long
    Dim FoundSelectedFile As Boolean
    
    ' Set MousePointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    FoundSelectedFile = False
    frmBranch.txtFile.Text = ""
    
    ' Set On Error routine
    On Error Resume Next
    
    ' Iterate through each item in the ListView Control
    For FileCount = 1 To ListView1.ListItems.Count
        
        ' File is selected
        If ListView1.ListItems(FileCount).Selected = True Then
            FoundSelectedFile = True
            
            ' Initialize the Branch dialog
            frmBranch.txtFile.Text = frmBranch.txtFile.Text + ListView1.ListItems(FileCount).Text + ", "
        End If
    Next
    If Right(frmBranch.txtFile.Text, 2) = ", " Then frmBranch.txtFile.Text = Left(frmBranch.txtFile.Text, Len(frmBranch.txtFile.Text) - 2)
    
    If Not FoundSelectedFile Then
    
        ' Instantiate the selected file
        Set objVSSProject = objVSSDatabase.VSSItem(ItemToBranch, False)
        FoundSelectedFile = True
           
        ' Initialize the Branch dialog
        frmBranch.txtFile.Text = frmBranch.txtFile.Text + objVSSProject.Name
        
    End If
    
    ' If there are files to be Branched then show the Branch Dialog
    If FoundSelectedFile Then
        frmBranch.Show 1
    Else
        Response = MsgBox("No file(s) selected", vbExclamation, AppTitle)
    End If
    
    ' Update GUI and Set Mousepointer
    RefreshFileList
    MousePointer = vbNormal
    
End Sub

Public Sub MoveItem()

    Dim objMoveItem As VSSItem
    
    ' Set Mousepointer
    MousePointer = vbHourglass

    ' Moving a file
    If Selected = "Listview1" Then
        
        'Instantiate the file
        Set objMoveItem = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
        
        ' Initialize Move dialog
        frmMove.Caption = "Move file " + objMoveItem.Spec
    
    ' Moving a project
    Else
    
        ' Check if user has selected the root project and set the VSSItem as appropriate
        If TreeView1.SelectedItem.Key <> "$" Then
            Set objMoveItem = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key, False)
        Else
            Set objMoveItem = objVSSDatabase.VSSItem(TreeView1.SelectedItem.Key + "/", False)
        End If
        
        ' Initialize Move dialog
        frmMove.Caption = "Move project " + objMoveItem.Spec
    End If
    
    ' Open the move dialog
    frmMove.Show 1
    
    ' Set Mousepointer
    MousePointer = vbNormal
    
End Sub

Public Sub ChangePassword()

    ' Show the Change Password Form
    frmChangePassword.Show 1

End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Show the popup menu
    If Button = vbRightButton Then
        frmMain.PopupMenu frmMain.mnuLeftPanePopupMenu
    End If
    
End Sub

' Resizes the information displayed above the List and Treeview controls
' to simulate a "Statusbar" like resize

Public Sub ResizeHeader()
        
    ListView1.Left = TreeView1.Width + PaneSeperator.Width
    ListView1.Width = frmMain.Width - TreeView1.Width + PaneSeperator.Width
    PaneSepLeft.X1 = PaneSeperator.Left - 30
    PaneSepLeft.X2 = PaneSepLeft.X1
    LinePaneSepRight.X1 = PaneSeperator.Left + PaneSeperator.Width
    LinePaneSepRight.X2 = LinePaneSepRight.X1
    LineProjectsBorderTop.X2 = PaneSepLeft.X1
    LineProjectsBorderTop.X2 = PaneSepLeft.X1
    linePaneSepLeft.X2 = PaneSepLeft.X1
    lineContentsBorderTop.X1 = LinePaneSepRight.X1
    lineContentsBorderBottom.X1 = LinePaneSepRight.X1
    lblContents.Left = LinePaneSepRight.X1 + 45
    lblAllProjects.Width = LineProjectsBorderTop.X2 - LineProjectsBorderTop.X1 - 100
    lineContentsRight.X1 = ListView1.Left + ListView1.Width / 2 - 20
    lineContentsRight.X2 = lineContentsRight.X1
    lineContentsBorderTop.X2 = lineContentsRight.X1
    lineContentsBorderBottom.X2 = lineContentsRight.X1
    lineWorkingFolderLeft.X1 = lineContentsRight.X1 + 60
    lineWorkingFolderLeft.X2 = lineWorkingFolderLeft.X1
    lineWorkingFolderBorder.X1 = lineWorkingFolderLeft.X1
    lineWorkingFolderBottom.X1 = lineWorkingFolderLeft.X1
    lineWorkingFolderBorder.X2 = frmMain.Width - 165
    lineWorkingFolderBottom.X2 = frmMain.Width - 165
    lblWorkingFolder.Left = lineWorkingFolderLeft.X1 + 45
    lblContents.Width = lineContentsBorderTop.X2 - lineContentsBorderTop.X1 - 100
    lineWorkingFolderRight.X1 = lineWorkingFolderBorder.X2
    lineWorkingFolderRight.X2 = lineWorkingFolderRight.X1
    lblWorkingFolder.Width = lineWorkingFolderBorder.X2 - lineWorkingFolderBorder.X1 - 100
    
End Sub

' Opens the Options dialog

Public Sub Options()

    frmOptions.Show 1

End Sub

Private Sub txtResults_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' Set on error routine
    On Error Resume Next
    
    If TypeOf Source Is PictureBox And Source.Name = "PaneSeperatorBottom" Then
        Source.Top = txtResults.Top + Y
        txtResults.Top = Source.Top + 80
        txtResults.Height = StatusBar1.Top - txtResults.Top
        ListView1.Height = txtResults.Top - ListView1.Top - 80
        TreeView1.Height = ListView1.Height
    ElseIf TypeOf Source Is ListView Then
        Call EndShareDrag(Source, X, Y)
    End If
    
End Sub

' Enables Menu, Toolbar and PopupMenu Items based on
' display of File List

Public Sub EnableFileControls(Enable As Boolean)

    mnuViewFile.Enabled = Enable
    mnuEditFile.Enabled = Enable
    mnuPopView.Enabled = Enable
    mnuPopEdit.Enabled = Enable
    mnuHistory.Enabled = Enable
    mnuShowDifferences.Enabled = Enable
    mnuPopHistory.Enabled = Enable
    mnuPopShowDifferences.Enabled = Enable
    mnuPopRename.Enabled = Enable
    mnuPopProperties.Enabled = Enable
    mnuProperties.Enabled = Enable
    mnuRename.Enabled = Enable
    Toolbar1.Buttons.Item("ViewFile").Enabled = Enable
    Toolbar1.Buttons.Item("EditFile").Enabled = Enable
    Toolbar1.Buttons.Item("ShowDifferences").Enabled = Enable
    Toolbar1.Buttons.Item("Properties").Enabled = Enable
    Toolbar1.Buttons.Item("History").Enabled = Enable
        
End Sub

' This routine is called when the user has asked to edit a file
' from the Menu using the OCX Control. It checks to see if the file is
' already checked out and, if it is opens it in the viewer. Otherwise
' it Checks the file out to the Working Dir and then opens it in the viewer.

Public Sub EditFileOCX()

    Dim RetVal As Long
    Dim Response As Long
    Dim objVSSFile As VSSItem
    Dim WorkingFolder As String

    ' Set Mousepointer
    MousePointer = vbHourglass
    
    ' Initialize variables
    Response = vbYes
   
    ' Set On Error routine
    On Error GoTo ErrHandler
    
    ' Set current VSS Item to selected file in ListView control
    Set objVSSFile = objVSSDatabase.VSSItem(ListView1.SelectedItem.Key, False)
    If Err = 0 Then
        
        ' Warn user if file is binary
        If objVSSFile.Binary = True Then Response = MsgBox("The file " + objVSSFile.Name + " is binary. Continue?", vbYesNo + vbQuestion)

        ' User wants to edit file
        If Response = vbYes Then
        
            ' File is already checked out by user
            If objVSSFile.IsCheckedOut = VSSFILE_CHECKEDOUT_ME Then
                For Each objVSSCheckout In objVSSFile.Checkouts
                    If objVSSCheckout.UserName = UserName Then
                        WorkingFolder = objVSSCheckout.LocalSpec
                        Exit For
                    End If
                Next
                                
                ' Display writeable copy in NotePad
                If Right(WorkingFolder, 1) <> "\" Then WorkingFolder = WorkingFolder + "\"
                
                ' Edit the file
                Viewer1.EditFile (WorkingFolder + objVSSFile.Name)
            
            ' Check the file out to the working directory
            Else
                FileDate = Date + Time
                
                ' Check for Working Directory
                On Error Resume Next
                WorkingFolder = objVSSFile.Parent.LocalSpec
                If WorkingFolder = "" Then
                    Response = MsgBox("You must set a working folder for this file before it can be edited." + vbCrLf + "Would you like to set one now?", vbExclamation + vbYesNo, AppTitle)
                    If Response = vbYes Then Call SetWorkingFolder
                Else
                
                    ' CheckOut the file and edit in NotePad
                    On Error GoTo ErrHandler
                    If Right(WorkingFolder, 1) <> "\" Then WorkingFolder = WorkingFolder + "\"
                    objVSSFile.CheckOut Comment:="", Local:=WorkingFolder + objVSSFile.Name, iFlags:=EOLFlag + VSSFLAG_CHKEXCLUSIVENO
                    frmMain.txtResults.Text = frmMain.txtResults.Text + objVSSFile.Name + " checked out" + vbCrLf
                    Call ShowResults
                    
                    ' Display writeable copy in NotePad
                    Viewer1.EditFile (WorkingFolder + objVSSFile.Name)
                End If
            End If
        End If
    End If
    
    ' Check for errors
    If Err <> 0 Then
ErrHandler:
        
        ' No working folder set
        If Err <> -2147166539 Then
            Response = MsgBox("Unable to edit file!" + vbCrLf + Err.Description, vbExclamation, AppTitle)
        End If
    End If
    
    '  Update GUI and Set Mousepointer
    RefreshFileList
    MousePointer = vbNormal
    
End Sub

' Ends the drag and drop share operation from the ListView control

Public Sub EndShareDrag(Source As Control, X As Single, Y As Single)
    
    If TypeOf Source Is ListView Or TypeOf Source Is TreeView Then
        ListView1.Drag vbEndDrag
        Dragging = EndDrag
        WhichButton = -99
        TreeView1.DropHighlight = Nothing
    End If

End Sub

