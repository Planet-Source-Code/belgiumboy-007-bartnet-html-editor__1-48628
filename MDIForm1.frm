VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm frmMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "BartNet HTML Editor"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   10770
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ilsToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   22
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            Object.Tag             =   ""
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   ""
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   ""
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            Object.Tag             =   ""
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            Object.Tag             =   ""
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   ""
            ImageKey        =   "Stop"
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   ""
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Preview"
            Object.Tag             =   ""
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            Object.Tag             =   ""
            ImageKey        =   "Options"
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            Object.Tag             =   ""
            ImageKey        =   "About"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   6000
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10770
      TabIndex        =   4
      Top             =   7935
      Width           =   10770
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Text            =   "fjlsdmq jfklm dqjklf mdsjqkflmsd"
         Top             =   105
         Width           =   6135
      End
      Begin VB.PictureBox Picture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   50
         Picture         =   "MDIForm1.frx":5C12
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   105
         Width           =   240
      End
      Begin ComctlLib.ProgressBar pStatus 
         Height          =   255
         Left            =   7200
         TabIndex        =   5
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
         Min             =   1e-4
         Max             =   1e12
      End
      Begin ComctlLib.StatusBar StatusBar 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   661
         Style           =   1
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   2
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Status"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   176
               MinWidth        =   176
               Key             =   "Progress"
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picMenus 
      Align           =   1  'Align Top
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   10710
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   10770
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   735
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   615
         ExtentX         =   1085
         ExtentY         =   1296
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   10080
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   26
         Left            =   6960
         Picture         =   "MDIForm1.frx":5F54
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   25
         Left            =   6960
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgEdit 
         Height          =   240
         Left            =   9600
         Picture         =   "MDIForm1.frx":6296
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgPreview 
         Height          =   240
         Left            =   9360
         Picture         =   "MDIForm1.frx":65D8
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   33
         Left            =   7680
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   34
         Left            =   7680
         Picture         =   "MDIForm1.frx":691A
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   38
         Left            =   8040
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   35
         Left            =   8040
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   45
         Left            =   8400
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   44
         Left            =   8400
         Picture         =   "MDIForm1.frx":6C5C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   43
         Left            =   8400
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   42
         Left            =   8400
         Picture         =   "MDIForm1.frx":6F9E
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   32
         Left            =   7680
         Picture         =   "MDIForm1.frx":72E0
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   31
         Left            =   7680
         Picture         =   "MDIForm1.frx":7622
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   30
         Left            =   7680
         Picture         =   "MDIForm1.frx":7964
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   17
         Left            =   6600
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   12
         Left            =   6600
         Picture         =   "MDIForm1.frx":7CA6
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   11
         Left            =   6600
         Picture         =   "MDIForm1.frx":7FE8
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   46
         Left            =   8400
         Picture         =   "MDIForm1.frx":832A
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   29
         Left            =   7680
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   28
         Left            =   7320
         Picture         =   "MDIForm1.frx":866C
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   27
         Left            =   7320
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   40
         Left            =   8040
         Picture         =   "MDIForm1.frx":89AE
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   39
         Left            =   8040
         Picture         =   "MDIForm1.frx":8CF0
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   23
         Left            =   6960
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   22
         Left            =   6960
         Picture         =   "MDIForm1.frx":9032
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   36
         Left            =   8040
         Picture         =   "MDIForm1.frx":9374
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   37
         Left            =   8040
         Picture         =   "MDIForm1.frx":96B6
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   24
         Left            =   6960
         Picture         =   "MDIForm1.frx":99F8
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   21
         Left            =   6960
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   10
         Left            =   6600
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   20
         Left            =   6600
         Picture         =   "MDIForm1.frx":9D3A
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   19
         Left            =   6600
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   18
         Left            =   6600
         Picture         =   "MDIForm1.frx":A07C
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   13
         Left            =   6600
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   14
         Left            =   6600
         Picture         =   "MDIForm1.frx":A3BE
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   16
         Left            =   6600
         Picture         =   "MDIForm1.frx":A700
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   15
         Left            =   6600
         Picture         =   "MDIForm1.frx":AA42
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   4
         Left            =   6240
         Picture         =   "MDIForm1.frx":AD84
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   8
         Left            =   6240
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   9
         Left            =   6240
         Picture         =   "MDIForm1.frx":B0C6
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   5
         Left            =   6240
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   7
         Left            =   6240
         Picture         =   "MDIForm1.frx":B408
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   6
         Left            =   6240
         Picture         =   "MDIForm1.frx":B74A
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   3
         Left            =   6240
         Picture         =   "MDIForm1.frx":BA8C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   2
         Left            =   6240
         Picture         =   "MDIForm1.frx":BDCE
         Top             =   840
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   1
         Left            =   6240
         Top             =   480
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   6240
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Left            =   10200
         Picture         =   "MDIForm1.frx":C110
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgUncheck 
         Height          =   240
         Left            =   10200
         Picture         =   "MDIForm1.frx":C452
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   41
         Left            =   8400
         Top             =   120
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.ImageList ilsToolbar 
      Left            =   5280
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":C794
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":CAE6
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":CE38
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D18A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D4DC
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D82E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":DB80
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":DED2
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":E224
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":E576
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":E8C8
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":EC1A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":EF6C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":F2BE
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":F610
            Key             =   "About"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":F962
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":FCB4
            Key             =   "Edit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuNothing1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuNothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
      End
      Begin VB.Menu mnuNothing3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuNothing4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuNothing5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu mnuNothing6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "StatusBar"
      End
      Begin VB.Menu mnuNothing11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuCascadeWindows 
         Caption         =   "Cascade Windows"
      End
      Begin VB.Menu mnuTileWindowsHorizontally 
         Caption         =   "Tile Windows Horizontally"
      End
      Begin VB.Menu mnuTileWindowsVertically 
         Caption         =   "Tile Windows Vertically"
      End
      Begin VB.Menu mnuNothing10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToWindow 
         Caption         =   "GoTo Window..."
      End
   End
   Begin VB.Menu mnuBrowserControls 
      Caption         =   "&Browser Controls"
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "Forward"
      End
      Begin VB.Menu mnuNothing9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCheckForUpdate 
         Caption         =   "Check For Update"
      End
      Begin VB.Menu mnuNothing7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisitBartNetOnline 
         Caption         =   "Visit BartNet Online"
      End
      Begin VB.Menu mnuNothing8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Dim pnt As PaintEffects
Dim MyFont As Long
Dim OldFont As Long
Dim wlOldProc As Long
Public LastIndex As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

Public Sub UpdateView()
    If Me.ActiveForm.picPreview.Visible = True Then
        Caps(26) = "Edit"
        Set frmMDI.img(26).Picture = frmMDI.imgEdit.Picture
        frmMDI.Toolbar.Buttons("Preview").ToolTipText = "Edit"
        frmMDI.Toolbar.Buttons("Preview").Image = "Edit"
    Else
        Caps(26) = "Preview"
        Set frmMDI.img(26).Picture = frmMDI.imgPreview.Picture
        frmMDI.Toolbar.Buttons("Preview").ToolTipText = "Preview"
        frmMDI.Toolbar.Buttons("Preview").Image = "Preview"
    End If
    
    pStatus.Visible = False
    txtStatus.Text = "Done"
End Sub

Private Sub MDIForm_Load()
    pStatus.Visible = False
    txtStatus.Text = "Done"
    
    Set pnt = New PaintEffects
    
    Caps(2) = "New..."
    Caps(3) = "Open..."
    Caps(4) = "Close"
    Caps(5) = ""
    Caps(6) = "Save"
    Caps(7) = "Save As..."
    Caps(8) = ""
    Caps(9) = "Exit"
    
    Caps(11) = "Undo"
    Caps(12) = "Redo"
    Caps(13) = ""
    Caps(14) = "Cut"
    Caps(15) = "Copy"
    Caps(16) = "Paste"
    Caps(17) = ""
    Caps(18) = "Select All"
    Caps(19) = ""
    Caps(20) = "Find..."
    
    Caps(22) = "Toolbar"
    Caps(23) = ""
    Caps(24) = "StatusBar"
    Caps(25) = ""
    Caps(26) = "Preview"
    
    Caps(28) = "Options..."
    
    Caps(30) = "Cascade Windows"
    Caps(31) = "Tile Windows Horizontally"
    Caps(32) = "Tile Windows Vertically"
    Caps(33) = ""
    Caps(34) = "GoTo Window..."
    
    Caps(36) = "Back"
    Caps(37) = "Forward"
    Caps(38) = ""
    Caps(39) = "Stop"
    Caps(40) = "Refresh"
    
    Caps(42) = "Check For Update..."
    Caps(43) = ""
    Caps(44) = "Visit BartNet Online..."
    Caps(45) = ""
    Caps(46) = "About..."
    
    If wlOldProc <> 0 Then Exit Sub
    
    Dim i As Integer
    
    MenuItems.MenuForm = Me

    MenuItems.SubMenu = 0
    For i = 0 To 11
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 1
    For i = 0 To 12
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 2
    For i = 0 To 7
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 3
    For i = 0 To 3
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next

    MenuItems.SubMenu = 4
    For i = 0 To 7
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 5
    For i = 0 To 7
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 6
    For i = 0 To 7
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
        
    wlOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf OwnMenuProc)
    
    Me.Caption = ProgName
End Sub

Public Function MsgProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim IsSep As Boolean
    Dim hBr As Long, hOldBr As Long
    Dim hPEN As Long, hOldPen As Long
    Dim lTextColor As Long
    Dim iRectOffset As Integer
    
    If wMsg = WM_DRAWITEM Then
        If wParam = 0 Then
            Call CopyMem(DrawInfo, ByVal lParam, LenB(DrawInfo))
            IsSep = IsSeparator(DrawInfo.itemID)
            
            MyFont = SendMessage(Me.hwnd, WM_GETFONT, 0&, 0&)
            OldFont = SelectObject(DrawInfo.hdc, MyFont)
            If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                hBr = CreateSolidBrush( _
                GetSysColor(COLOR_HIGHLIGHT))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            Else
                hBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
                hPEN = GetPen(1, GetSysColor(COLOR_MENU))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            End If
            QuickGDI.TargethDC = DrawInfo.hdc
            

            hOldBr = SelectObject(DrawInfo.hdc, hBr)
            hOldPen = SelectObject(DrawInfo.hdc, hPEN)
            With DrawInfo.rcItem
                If (DrawInfo.itemState And ODS_SELECTED) <> ODS_SELECTED Then
                    QuickGDI.DrawRect .Left, .Top, 22, .Bottom
                End If
                
                iRectOffset = IIf(img(DrawInfo.itemID).Picture.Handle <> 0 _
                    , 23, 0)
                If Not IsSep Then
                    
                    QuickGDI.DrawRect .Left + iRectOffset, .Top, .Right, .Bottom
                    
                    DrawFilledRect DrawInfo.hdc, .Left, .Top, .Right, .Bottom, vbWhite
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        DrawFilledRect1 .Left, .Top, .Right, .Bottom
                    End If
                    
                    'Print the item's text
                    '(held in the Caps() array)
                    hPrint .Left + 30, .Top + 3, Caps(DrawInfo.itemID), lTextColor
                End If
            End With
            Call SelectObject(DrawInfo.hdc, hOldBr)
            Call SelectObject(DrawInfo.hdc, hOldPen)
            Call DeleteObject(hBr)
            Call DeleteObject(hPEN)
            With DrawInfo
                If img(DrawInfo.itemID).Picture.Handle <> 0 Then
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        Dim i As Long
                        Dim e As Long
                        Dim a As Long
                        
                        Picture2.Cls
                        
                        pnt.PaintTransparentStdPic Picture2.hdc, 0, 0, _
                            16, 16, img(DrawInfo.itemID).Picture, 0, 0, vbMagenta
                        
                        For i = 0 To 16
                            For e = 0 To 16
                                a = GetPixel(Picture2.hdc, i, e)
                                If a <> vbMagenta Then SetPixel Picture2.hdc, i, e, RGB(158, 158, 165)
                            Next
                        Next
                    
                        Picture2.Refresh
                        
                                pnt.PaintTransparentDC .hdc, _
                            5, .rcItem.Top + 6, _
                            16, 16, Picture2.hdc, _
                            0, 0, vbMagenta
                                                    
                        pnt.PaintTransparentStdPic .hdc, _
                            3, .rcItem.Top + 4, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    Else
                        DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                        
                        pnt.PaintTransparentStdPic .hdc, _
                            4, .rcItem.Top + 5, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    End If
                    
                End If
                If IsSep Then
                    Dim pt As POINTAPI
                    DrawFilledRect .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, vbWhite
                    DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                    MoveToEx .hdc, .rcItem.Left + 25, .rcItem.Top + 2, pt
                    LineTo .hdc, .rcItem.Right, .rcItem.Top + 2
                End If
            End With
        End If
        MsgProc = False
        Exit Function
        
    ElseIf wMsg = WM_MEASUREITEM Then
        Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
        IsSep = IsSeparator(MeasureInfo.itemID)
        MeasureInfo.itemWidth = 150
        MeasureInfo.itemHeight = IIf(IsSep, 5, 22)
        Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
        MsgProc = False
        Exit Function
    ElseIf wMsg = WM_MENUSELECT Then
        
    End If
    
    MsgProc = CallWindowProc(wlOldProc, hwnd, wMsg, wParam, lParam)
End Function

Public Function IsSeparator(ByVal IID As Integer) As Boolean
    Dim mii As MENUITEMINFO
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_TYPE
    mii.wID = IID
    GetMenuItemInfo GetMenu(hwnd), IID, False, mii
    IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Timer.Enabled = True
    Else
        Unload frmAbout
        Unload frmExit
        Unload frmFind
        Unload frmNew
        Unload frmOpen
        Unload frmOptions
        Unload frmSplash
        Unload frmStartUp
        Unload frmUpdate
        Unload frmWindows
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then frmFind.Visible = False

    If Me.WindowState <> vbMinimized Then
        If Me.Width >= 7000 Then
            txtStatus.Width = Me.ScaleWidth - 50 - 1500
            pStatus.Left = txtStatus.Left + txtStatus.Width - 250
            StatusBar.Width = Me.Width
        Else
            Me.Width = 7000
        End If
        
        If Me.Height >= 5000 Then
        
        Else
            Me.Height = 5000
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If wlOldProc <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, wlOldProc
    End If
    
    Set pnt = Nothing

    DeleteObject MyFont
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBack_Click()
On Error Resume Next
    If Me.ActiveForm.picPreview.Visible = True Then Me.ActiveForm.WebBrowser.GoBack
End Sub

Private Sub mnuCascadeWindows_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCheckForUpdate_Click()
    frmUpdate.Show vbModal, Me
End Sub

Public Sub mnuClose_Click()
On Error Resume Next
    Dim a As Integer
    
    If Me.ActiveForm.Saved = False Then
        If Me.ActiveForm.SavedBefore = True Then
            a = MsgBox("Do you want to save the changes you made to " & Me.ActiveForm.Caption & " ?", vbYesNoCancel + vbQuestion, ProgName)
        
            Select Case a
                Case vbYes
                    mnuSave_Click
                    Unload Me.ActiveForm
                Case vbNo
                    Unload Me.ActiveForm
                Case vbCancel
                    
            End Select
        Else
            a = MsgBox("Do you want to save the changes you made to " & Me.ActiveForm.Caption & " ?", vbYesNoCancel + vbQuestion, ProgName)
        
            Select Case a
                Case vbYes
                    Dim fso As New FileSystemObject
                    Dim strm As TextStream
                    Dim File As File
                    Dim Continue As Boolean
    
    On Error GoTo err
                    Set File = fso.GetFile(App.Path & "\Recent")

                    CommonDialog.Filter = "Internet Documents (HTML)|*.htm;*.html"
                    CommonDialog.FileName = Me.ActiveForm.Caption
    On Error GoTo CommErr
                    CommonDialog.ShowSave
    On Error Resume Next
                    Set strm = fso.CreateTextFile(CommonDialog.FileName, True)
                    strm.Write Me.ActiveForm.RichTextBox.Text
                    strm.Close
       
                    Set File = fso.GetFile(CommonDialog.FileName)
        
                    Continue = True

                    Set strm = fso.OpenTextFile(App.Path & "\Recent", ForReading)

                    Do Until strm.AtEndOfStream
                        If strm.ReadLine = File.Name Then Continue = False
                    Loop
                    strm.Close


    On Error Resume Next
                    If Continue = True Then
                        Dim strm2 As TextStream
                        
                        fso.CopyFile App.Path & "\Recent", fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", True
                        
                        Set strm2 = fso.OpenTextFile(fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", ForReading)
                        Set strm = fso.OpenTextFile(App.Path & "\Recent", ForWriting)
                        strm.WriteLine File.Name
                        strm.WriteLine Mid(File.Path, 1, Len(File.Path) - Len(File.Name))
                        Do Until strm2.AtEndOfStream
                            strm.WriteLine strm2.ReadLine
                        Loop
                        strm.Close
                        strm2.Close
                    End If

                    Unload Me.ActiveForm
                Case vbNo
                    Unload Me.ActiveForm
                Case vbCancel
                    
            End Select
        End If
    Else
        Unload Me.ActiveForm
    End If
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(App.Path & "\Recent")
    strm.Close
    
    Set File = fso.GetFile(App.Path & "\Recent")
    
    File.Attributes = Hidden + System
    
    Set File = Nothing
    
    mnuClose_Click
    
CommErr:
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.RichTextBox.SelText, 1
End Sub

Private Sub mnuCut_Click()
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.RichTextBox.SelText, 1
    
    Me.ActiveForm.RichTextBox.SelText = ""
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    If FormsCount = 0 Then
        Unload Me
    Else
        frmExit.Show vbModal, Me
    End If
End Sub

Private Sub mnuFind_Click()
    If FormsCount <> 0 Then frmFind.Show vbModeless, Me
End Sub

Private Sub mnuForward_Click()
On Error Resume Next
    If Me.ActiveForm.picPreview.Visible = True Then Me.ActiveForm.WebBrowser.GoForward
End Sub

Private Sub mnuGoToWindow_Click()
On Error Resume Next
    frmWindows.Show vbModal, Me
End Sub

Private Sub mnuNew_Click()
    frmNew.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
    frmOpen.Show vbModal, Me
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuPaste_Click()
On Error Resume Next
    Me.ActiveForm.RichTextBox.SelText = Clipboard.GetText(1)
End Sub

Private Sub mnuPreview_Click()
On Error Resume Next
    If FormsCount = 0 Then Exit Sub

    If img(26).Picture = imgEdit Then
        Caps(26) = "Preview"
        Set img(26).Picture = imgPreview.Picture
        Toolbar.Buttons("Preview").ToolTipText = "Preview"
        Toolbar.Buttons("Preview").Image = "Preview"
        Me.ActiveForm.picPreview.Visible = False
        Me.ActiveForm.picEdit.Visible = True
    Else
        If Me.ActiveForm.Saved = True Then
            Me.ActiveForm.WebBrowser.Navigate Me.ActiveForm.SavedBeforePath
            txtStatus.Text = "Opening page " & Me.ActiveForm.SavedBeforePath
        Else
            If Me.ActiveForm.SavedBefore = True Then
                mnuSave_Click
                Me.ActiveForm.WebBrowser.Navigate Me.ActiveForm.SavedBeforePath
                txtStatus.Text = "Opening page " & Me.ActiveForm.SavedBeforePath
            Else
                Dim fso As New FileSystemObject
                Dim strm As TextStream
                
                Set strm = fso.CreateTextFile(fso.GetSpecialFolder(TemporaryFolder) & "\" & Me.ActiveForm.Caption & "bntmp.html", True)
                strm.Write Me.ActiveForm.RichTextBox.Text
                strm.Close
                
                Me.ActiveForm.WebBrowser.Navigate fso.GetSpecialFolder(TemporaryFolder) & "\" & Me.ActiveForm.Caption & "bntmp.html"
                txtStatus.Text = "Opening page " & fso.GetSpecialFolder(TemporaryFolder) & "\" & Me.ActiveForm.Caption & "bntmp.html"
            End If
        End If
        
        Caps(26) = "Edit"
        Set img(26).Picture = imgEdit.Picture
        Toolbar.Buttons("Preview").ToolTipText = "Edit"
        Toolbar.Buttons("Preview").Image = "Edit"
        Me.ActiveForm.picPreview.Visible = True
        Me.ActiveForm.picEdit.Visible = False
    End If
End Sub

Private Sub mnuRedo_Click()
On Error Resume Next
    Me.ActiveForm.Redo
End Sub

Private Sub mnuRefresh_Click()
On Error Resume Next
    If Me.ActiveForm.picPreview.Visible = True Then Me.ActiveForm.WebBrowser.Refresh
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
    If Me.ActiveForm.Saved = True Then Exit Sub

    If Me.ActiveForm.SavedBefore = True Then
        Dim fso As New FileSystemObject
        Dim strm As TextStream
On Error GoTo err
        Set strm = fso.OpenTextFile(Me.ActiveForm.SavedBeforePath, ForWriting)
        strm.Write Me.ActiveForm.RichTextBox.Text
        strm.Close
        
        Me.ActiveForm.Saved = True
    Else
        mnuSaveAs_Click
    End If
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(Me.ActiveForm.SavedBeforePath, True)
    strm.Write Me.ActiveForm.RichTextBox.Text
    strm.Close
    
    Me.ActiveForm.Saved = True
End Sub

Private Sub mnuSaveAs_Click()
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim File As File
    Dim Continue As Boolean
    
On Error GoTo err
    Set File = fso.GetFile(App.Path & "\Recent")

    CommonDialog.Filter = "Internet Documents (HTML)|*.htm;*.html"
    CommonDialog.FileName = Me.ActiveForm.Caption
On Error GoTo CommErr
    CommonDialog.ShowSave
On Error Resume Next
    Set strm = fso.CreateTextFile(CommonDialog.FileName, True)
    strm.Write Me.ActiveForm.RichTextBox.Text
    strm.Close
       
    Set File = fso.GetFile(CommonDialog.FileName)
        
    Continue = True

    Set strm = fso.OpenTextFile(App.Path & "\Recent", ForReading)

    Do Until strm.AtEndOfStream
        If strm.ReadLine = File.Name Then Continue = False
    Loop
    strm.Close
        
On Error Resume Next
    If Continue = True Then
        Dim strm2 As TextStream
        
        fso.CopyFile App.Path & "\Recent", fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", True
        
        Set strm2 = fso.OpenTextFile(fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", ForReading)
        Set strm = fso.OpenTextFile(App.Path & "\Recent", ForWriting)
        strm.WriteLine File.Name
        strm.WriteLine Mid(File.Path, 1, Len(File.Path) - Len(File.Name))
        Do Until strm2.AtEndOfStream
            strm.WriteLine strm2.ReadLine
        Loop
        strm.Close
        strm2.Close
    End If
    
    Me.ActiveForm.Saved = True
    Me.ActiveForm.SavedBefore = True
    Me.ActiveForm.SavedBeforePath = File.Path
    Me.ActiveForm.Caption = File.Name

    Exit Sub
  
err:
    Set strm = fso.CreateTextFile(App.Path & "\Recent")
    strm.Close
    
    Set File = fso.GetFile(App.Path & "\Recent")
    
    File.Attributes = Hidden + System
    
    Set File = Nothing
    
    mnuSaveAs_Click
    
    Exit Sub
    
CommErr:
End Sub

Private Sub mnuSelectAll_Click()
On Error Resume Next
    Me.ActiveForm.RichTextBox.SelStart = 0
    Me.ActiveForm.RichTextBox.SelLength = Len(Me.ActiveForm.RichTextBox.Text)
End Sub

Private Sub mnuStatusBar_Click()
    If img(24).Picture = imgCheck.Picture Then
        Set img(24).Picture = imgUncheck.Picture
        picStatus.Visible = False
        MDIForm_Resize
    Else
        Set img(24).Picture = imgCheck.Picture
        picStatus.Visible = True
        MDIForm_Resize
    End If
End Sub

Private Sub mnuStop_Click()
On Error Resume Next
    If Me.ActiveForm.picPreview.Visible = True Then Me.ActiveForm.WebBrowser.Stop
End Sub

Private Sub mnuTileWindowsHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileWindowsVertically_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuToolbar_Click()
    If img(22).Picture = imgCheck.Picture Then
        Set img(22).Picture = imgUncheck.Picture
        Toolbar.Visible = False
        MDIForm_Resize
    Else
        Set img(22).Picture = imgCheck.Picture
        Toolbar.Visible = True
        MDIForm_Resize
    End If
End Sub

Private Sub mnuUndo_Click()
On Error Resume Next
    Me.ActiveForm.Undo
End Sub

Private Sub mnuVisitBartNetOnline_Click()
    WebBrowser.Navigate2 "http://www.bartnet.be", , "_new"
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    mnuExit_Click
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "New"
            PageNumber = PageNumber + 1
    
            Dim tmp As String
    
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & "</body>" & vbCrLf & "</html>"
    
            LoadPage tmp, DocumentName & " " & PageNumber
        Case "Open"
            OpenFile
        Case "Save"
            mnuSave_Click
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Find"
            mnuFind_Click
        Case "Undo"
            mnuUndo_Click
        Case "Redo"
            mnuRedo_Click
        Case "Back"
            mnuBack_Click
        Case "Forward"
            mnuForward_Click
        Case "Stop"
            mnuStop_Click
        Case "Refresh"
            mnuRefresh_Click
        Case "Preview"
            mnuPreview_Click
        Case "Options"
            mnuOptions_Click
        Case "About"
            mnuAbout_Click
    End Select
End Sub

Private Sub OpenFile(Optional FirstTime As Boolean = True)
    Dim tmp As String
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim File As File
    Dim Continue As Boolean
    
    CommonDialog.Filter = "Internet Documents (HTML)|*.htm;*.html"
On Error GoTo CommErr
    If FirstTime = True Then CommonDialog.ShowOpen
On Error Resume Next
    Set strm = fso.OpenTextFile(CommonDialog.FileName, ForReading)
    tmp = strm.ReadAll
    strm.Close
    
    Set File = fso.GetFile(CommonDialog.FileName)
        
    Continue = True
On Error GoTo err
    Set strm = fso.OpenTextFile(App.Path & "\Recent", ForReading)
On Error Resume Next
    Do Until strm.AtEndOfStream
        If strm.ReadLine = File.Name Then Continue = False
    Loop
    strm.Close
        
On Error Resume Next
    If Continue = True Then
        Dim strm2 As TextStream
        
        fso.CopyFile App.Path & "\Recent", fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", True
        
        Set strm2 = fso.OpenTextFile(fso.GetSpecialFolder(TemporaryFolder) & "\tmpRecent.BartNet", ForReading)
        Set strm = fso.OpenTextFile(App.Path & "\Recent", ForWriting)
        strm.WriteLine File.Name
        strm.WriteLine Mid(File.Path, 1, Len(File.Path) - Len(File.Name))
        Do Until strm2.AtEndOfStream
            strm.WriteLine strm2.ReadLine
        Loop
        strm.Close
        strm2.Close
    End If
        
    LoadPage tmp, File.Name, False, CommonDialog.FileName
    
    Exit Sub
  
err:
    Set strm = fso.CreateTextFile(App.Path & "\Recent")
    strm.Close
    
    Set File = fso.GetFile(App.Path & "\Recent")
    
    File.Attributes = Hidden + System
    
    Set File = Nothing
    
    OpenFile False
    
    Exit Sub
    
CommErr:
End Sub
