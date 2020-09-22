VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   11280
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   12
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0000
            Key             =   "Fuschia"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":04D2
            Key             =   "Black"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":09A4
            Key             =   "Maroon"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0E76
            Key             =   "Green"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1348
            Key             =   "Olive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":181A
            Key             =   "Navy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1CEC
            Key             =   "Purple"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":21BE
            Key             =   "Teal"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2690
            Key             =   "Gray"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2B62
            Key             =   "Custom"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3034
            Key             =   "Silver"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3506
            Key             =   "White"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":39D8
            Key             =   "Red"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3EAA
            Key             =   "Lime"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":437C
            Key             =   "Yellow"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":484E
            Key             =   "Blue"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":4D20
            Key             =   "Aqua"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHidden 
      Height          =   255
      Left            =   12240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   4440
      ScaleHeight     =   1935
      ScaleWidth      =   3375
      TabIndex        =   26
      Top             =   5640
      Width           =   3375
      Begin VB.OptionButton optNewPage 
         Caption         =   "New Page"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optUntitled 
         Caption         =   "Untitled"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDocumentName 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1215
      ScaleWidth      =   2775
      TabIndex        =   22
      Top             =   6000
      Width           =   2775
      Begin VB.CheckBox ckStartUpScreen 
         Caption         =   "Show StartUp screen at startup"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox ckToolbar 
         Caption         =   "Show Toolbar by default"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox ckStatusbar 
         Caption         =   "Show Statusbar by default"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6000
      ScaleHeight     =   2535
      ScaleWidth      =   4695
      TabIndex        =   10
      Top             =   480
      Width           =   4695
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Tags Color :"
         Top             =   2080
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Comments Color :"
         Top             =   1720
         Width           =   1215
      End
      Begin VB.CheckBox ckSyntaxHighlighting 
         Caption         =   "Syntax Highlighting"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Default BackGround Color :"
         Top             =   880
         Width           =   1935
      End
      Begin VB.TextBox txtFontSize 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   500
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Default Font Size :"
         Top             =   520
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Default Font :"
         Top             =   160
         Width           =   975
      End
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   120
         Width           =   2415
      End
      Begin MSComctlLib.ImageCombo cboBackGroundColor 
         Height          =   330
         Left            =   2160
         TabIndex        =   16
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo cboComments 
         Height          =   330
         Left            =   2160
         TabIndex        =   20
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo cboTags 
         Height          =   330
         Left            =   2160
         TabIndex        =   21
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   2055
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   3360
      Width           =   4695
      Begin VB.TextBox txtSaveTo 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1680
         Width           =   3975
      End
      Begin VB.OptionButton optSaveAll 
         Caption         =   "Save all open documents"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optSave 
         Caption         =   "Save all documents which have been saved before"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton optDontSave 
         Caption         =   "Don't save any documents automatically"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.CheckBox ckShowExitScreen 
         Caption         =   "Show Exit Screen"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Save all unsaved documents to :"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   "General"
            Object.Tag             =   ""
            Object.ToolTipText     =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Default Document Name"
            Key             =   "Default Document Name"
            Object.Tag             =   "Default Document Name"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Appearance"
            Key             =   "Appearance"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Appearance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit Wizard"
            Key             =   "Exit Wizard"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Exit Wizard"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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

Private Changed As Boolean

Private Sub cboBackGroundColor_Click()
    If cboBackGroundColor.Text = "Custom" Then
        CommonDialog.ShowColor
        cboBackGroundColor.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub cboBackGroundColor_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboBackGroundColor.Text = "Custom" Then
        CommonDialog.ShowColor
        cboBackGroundColor.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub cboComments_Click()
    If cboComments.Text = "Custom" Then
        CommonDialog.ShowColor
        cboComments.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub cboComments_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboComments.Text = "Custom" Then
        CommonDialog.ShowColor
        cboComments.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub cboFont_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboFont_KeyUp(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub cboTags_Click()
    If cboTags.Text = "Custom" Then
        CommonDialog.ShowColor
        cboTags.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub cboTags_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboTags.Text = "Custom" Then
        CommonDialog.ShowColor
        cboTags.Tag = CommonDialog.Color
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub ckShowExitScreen_Click()
    If ckShowExitScreen.Value = vbChecked Then
        optDontSave.Enabled = False
        optSave.Enabled = False
        optSaveAll.Enabled = False
        Label1.Enabled = False
        txtSaveTo.Enabled = False
    Else
        optDontSave.Enabled = True
        optSave.Enabled = True
        optSaveAll.Enabled = True
        Label1.Enabled = True
        txtSaveTo.Enabled = True
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub ckStartUpScreen_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckStatusbar_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckSyntaxHighlighting_Click()
    If ckSyntaxHighlighting.Value = vbChecked Then
        Text4.Enabled = True
        Text5.Enabled = True
        cboComments.Enabled = True
        cboTags.Enabled = True
    Else
        Text4.Enabled = False
        Text5.Enabled = False
        cboComments.Enabled = False
        cboTags.Enabled = False
    End If

    cmdApply.Enabled = True
End Sub

Private Sub ckToolbar_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
On Error GoTo err
    Dim fso As New FileSystemObject
    Dim strm As TextStream
        
    Set strm = fso.OpenTextFile(App.Path & "\Settings", ForWriting)
    
    If ckStartUpScreen.Value = vbChecked Then
        strm.WriteLine "True"
    Else
        strm.WriteLine "False"
    End If
    
    If ckToolbar.Value = vbChecked Then
        strm.WriteLine "True"
    Else
        strm.WriteLine "False"
    End If
    
    If ckStatusbar.Value = vbChecked Then
        strm.WriteLine "True"
    Else
        strm.WriteLine "False"
    End If
    
    If optNewPage.Value = True Then
        strm.WriteLine "New Page"
    Else
        If optUntitled.Value = True Then
            strm.WriteLine "Untitled"
        Else
            strm.WriteLine txtDocumentName.Text
        End If
    End If
    
    strm.WriteLine cboFont.Text
    strm.WriteLine txtFontSize.Text
    
    If ckSyntaxHighlighting.Value = vbChecked Then
        If cboTags.Text = "Custom" Then
            strm.WriteLine cboTags.Tag
        Else
            strm.WriteLine GetColorFromString(cboTags.Text)
        End If
        
        If cboComments.Text = "Custom" Then
            strm.WriteLine cboComments.Tag
        Else
            strm.WriteLine GetColorFromString(cboComments.Text)
        End If
        
        strm.WriteLine "False"
    Else
        strm.WriteLine "8388608"
        strm.WriteLine "32768"
        strm.WriteLine "True"
    End If
    
    If cboBackGroundColor.Text = "Custom" Then
        strm.WriteLine cboBackGroundColor.Tag
    Else
        strm.WriteLine GetColorFromString(cboBackGroundColor.Text)
    End If
    
    If ckShowExitScreen.Value = vbChecked Then
        strm.WriteLine "0"
        strm.WriteLine "NONE"
    Else
        If optDontSave.Value = True Then
            strm.WriteLine "2"
            strm.WriteLine "NONE"
        Else
            If optSave.Value = True Then
                strm.WriteLine "1"
                strm.WriteLine "NONE"
            Else
                strm.WriteLine "3"
                strm.WriteLine txtSaveTo.Text
            End If
        End If
    End If
    
    strm.Close
    
    cmdApply.Enabled = False
    Changed = True
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(App.Path & "\Settings", True)

    strm.WriteLine "True"
    strm.WriteLine "True"
    strm.WriteLine "True"
    strm.WriteLine "New Page"
    strm.WriteLine "Courier New"
    strm.WriteLine "9"
    strm.WriteLine "16711680"
    strm.WriteLine "32768"
    strm.WriteLine "False"
    strm.WriteLine "16777215"
    strm.WriteLine "0"
    strm.WriteLine "NONE"
    
    strm.Close
    
    Dim File As File
    
    Set File = fso.GetFile(App.Path & "\Settings")
    
    File.Attributes = Hidden + System
    
    cmdApply_Click
End Sub

Private Sub cmdClose_Click()
    If Changed = True Then MsgBox "The settings that have been changed will function next time you run " & ProgName & ".", vbOKOnly + vbInformation, "Notice"
    
    Unload Me
End Sub

Private Sub Form_Load()
    For X = 1 To Screen.FontCount
        cboFont.AddItem Screen.Fonts(X)
    Next
    
    cboFont.RemoveItem 0
    
    Picture2.Top = 600
    Picture2.Left = 240
    Picture2.Visible = False
    
    Picture3.Top = 1080
    Picture3.Left = 1080
    
    Picture4.Top = 840
    Picture4.Left = 840
    Picture4.Visible = False
    
    Picture1.Top = 720
    Picture1.Left = 240
    Picture1.Visible = False
    
    Dim i As Long
    Dim colorUpper As String
    Dim colorProper As String
    
    Colors = Split(COLOR_LIST, ",")
    
    cboBackGroundColor.ImageList = ImageList
    cboComments.ImageList = ImageList
    cboTags.ImageList = ImageList
    
    cboBackGroundColor.ComboItems.Add , "Black", "Black", "Black"
    cboComments.ComboItems.Add , "Black", "Black", "Black"
    cboTags.ComboItems.Add , "Black", "Black", "Black"
    
    cboBackGroundColor.ComboItems.Add , "Maroon", "Maroon", "Maroon"
    cboComments.ComboItems.Add , "Maroon", "Maroon", "Maroon"
    cboTags.ComboItems.Add , "Maroon", "Maroon", "Maroon"
    
    cboBackGroundColor.ComboItems.Add , "Green", "Green", "Green"
    cboComments.ComboItems.Add , "Green", "Green", "Green"
    cboTags.ComboItems.Add , "Green", "Green", "Green"
    
    cboBackGroundColor.ComboItems.Add , "Olive", "Olive", "Olive"
    cboComments.ComboItems.Add , "Olive", "Olive", "Olive"
    cboTags.ComboItems.Add , "Olive", "Olive", "Olive"
    
    cboBackGroundColor.ComboItems.Add , "Navy", "Navy", "Navy"
    cboComments.ComboItems.Add , "Navy", "Navy", "Navy"
    cboTags.ComboItems.Add , "Navy", "Navy", "Navy"
    
    cboBackGroundColor.ComboItems.Add , "Purple", "Purple", "Purple"
    cboComments.ComboItems.Add , "Purple", "Purple", "Purple"
    cboTags.ComboItems.Add , "Purple", "Purple", "Purple"
    
    cboBackGroundColor.ComboItems.Add , "Teal", "Teal", "Teal"
    cboComments.ComboItems.Add , "Teal", "Teal", "Teal"
    cboTags.ComboItems.Add , "Teal", "Teal", "Teal"
    
    cboBackGroundColor.ComboItems.Add , "Gray", "Gray", "Gray"
    cboComments.ComboItems.Add , "Gray", "Gray", "Gray"
    cboTags.ComboItems.Add , "Gray", "Gray", "Gray"
    
    cboBackGroundColor.ComboItems.Add , "Silver", "Silver", "Silver"
    cboComments.ComboItems.Add , "Silver", "Silver", "Silver"
    cboTags.ComboItems.Add , "Silver", "Silver", "Silver"
    
    cboBackGroundColor.ComboItems.Add , "Red", "Red", "Red"
    cboComments.ComboItems.Add , "Red", "Red", "Red"
    cboTags.ComboItems.Add , "Red", "Red", "Red"
    
    cboBackGroundColor.ComboItems.Add , "Lime", "Lime", "Lime"
    cboComments.ComboItems.Add , "Lime", "Lime", "Lime"
    cboTags.ComboItems.Add , "Lime", "Lime", "Lime"
    
    cboBackGroundColor.ComboItems.Add , "Yellow", "Yellow", "Yellow"
    cboComments.ComboItems.Add , "Yellow", "Yellow", "Yellow"
    cboTags.ComboItems.Add , "Yellow", "Yellow", "Yellow"
    
    cboBackGroundColor.ComboItems.Add , "Blue", "Blue", "Blue"
    cboComments.ComboItems.Add , "Blue", "Blue", "Blue"
    cboTags.ComboItems.Add , "Blue", "Blue", "Blue"
    
    cboBackGroundColor.ComboItems.Add , "Fuschia", "Fuschia", "Fuschia"
    cboComments.ComboItems.Add , "Fuschia", "Fuschia", "Fuschia"
    cboTags.ComboItems.Add , "Fuschia", "Fuschia", "Fuschia"
    
    cboBackGroundColor.ComboItems.Add , "Aqua", "Aqua", "Aqua"
    cboComments.ComboItems.Add , "Aqua", "Aqua", "Aqua"
    cboTags.ComboItems.Add , "Aqua", "Aqua", "Aqua"
    
    cboBackGroundColor.ComboItems.Add , "White", "White", "White"
    cboComments.ComboItems.Add , "White", "White", "White"
    cboTags.ComboItems.Add , "White", "White", "White"
    
    cboBackGroundColor.ComboItems.Add , "Custom", "Custom", "Custom"
    cboComments.ComboItems.Add , "Custom", "Custom", "Custom"
    cboTags.ComboItems.Add , "Custom", "Custom", "Custom"
    
    cboBackGroundColor.ComboItems(1).Selected = True
    cboComments.ComboItems(1).Selected = True
    cboTags.ComboItems(1).Selected = True
    
    cboBackGroundColor.Refresh
    cboComments.Refresh
    cboTags.Refresh
    
    If StartUpScreenShow = True Then
        ckStartUpScreen.Value = vbChecked
    Else
        ckStartUpScreen.Value = vbUnchecked
    End If
    
    If ShowToolBar = True Then
        ckToolbar.Value = vbChecked
    Else
        ckToolbar.Value = vbUnchecked
    End If
    
    If ShowStatusBar = True Then
        ckStatusbar.Value = vbChecked
    Else
        ckStatusbar.Value = vbUnchecked
    End If

    If DocumentName = "New Page" Then
        optNewPage.Value = True
        txtDocumentName.Text = ""
        txtDocumentName.Enabled = False
    Else
        If DocumentName = "Untitled" Then
            optUntitled.Value = True
            txtDocumentName.Text = ""
            txtDocumentName.Enabled = False
        Else
            optOther.Value = True
            txtDocumentName.Text = DocumentName
            txtDocumentName.Enabled = True
        End If
    End If

    cboFont.Text = DefaultFontType
    txtFontSize.Text = DefaultFontSize

    Select Case DefaultBackGroundColor
        Case &H0&
            cboBackGroundColor.ComboItems.Item(1).Selected = True
        Case &H80&
            cboBackGroundColor.ComboItems.Item(2).Selected = True
        Case &H8000&
            cboBackGroundColor.ComboItems.Item(3).Selected = True
        Case &H8080&
            cboBackGroundColor.ComboItems.Item(4).Selected = True
        Case &H800000
            cboBackGroundColor.ComboItems.Item(5).Selected = True
        Case &H800080
            cboBackGroundColor.ComboItems.Item(6).Selected = True
        Case &H808000
            cboBackGroundColor.ComboItems.Item(7).Selected = True
        Case &H808080
            cboBackGroundColor.ComboItems.Item(8).Selected = True
        Case &HC0C0C0
            cboBackGroundColor.ComboItems.Item(9).Selected = True
        Case &HFF&
            cboBackGroundColor.ComboItems.Item(10).Selected = True
        Case &HFF00&
            cboBackGroundColor.ComboItems.Item(11).Selected = True
        Case &HFFFF&
            cboBackGroundColor.ComboItems.Item(12).Selected = True
        Case &HFF0000
            cboBackGroundColor.ComboItems.Item(13).Selected = True
        Case &HFF00FF
            cboBackGroundColor.ComboItems.Item(14).Selected = True
        Case &HFFFF00
            cboBackGroundColor.ComboItems.Item(15).Selected = True
        Case &HFFFFFF
            cboBackGroundColor.ComboItems.Item(16).Selected = True
        Case Else
            cboBackGroundColor.ComboItems.Item(17).Selected = True
            cboBackGroundColor.Tag = DefaultBackGroundColor
    End Select
    
    If DisableTagColoring = True Then
        ckSyntaxHighlighting.Value = vbUnchecked
        Text4.Enabled = False
        Text5.Enabled = False
        cboComments.Enabled = False
        cboTags.Enabled = False
    Else
        ckSyntaxHighlighting.Value = vbChecked
        Text4.Enabled = True
        Text5.Enabled = True
        cboComments.Enabled = True
        cboTags.Enabled = True
    End If
    
    Select Case HTML_Color
        Case &H0&
            cboTags.ComboItems.Item(1).Selected = True
        Case &H80&
            cboTags.ComboItems.Item(2).Selected = True
        Case &H8000&
            cboTags.ComboItems.Item(3).Selected = True
        Case &H8080&
            cboTags.ComboItems.Item(4).Selected = True
        Case &H800000
            cboTags.ComboItems.Item(5).Selected = True
        Case &H800080
            cboTags.ComboItems.Item(6).Selected = True
        Case &H808000
            cboTags.ComboItems.Item(7).Selected = True
        Case &H808080
            cboTags.ComboItems.Item(8).Selected = True
        Case &HC0C0C0
            cboTags.ComboItems.Item(9).Selected = True
        Case &HFF&
            cboTags.ComboItems.Item(10).Selected = True
        Case &HFF00&
            cboTags.ComboItems.Item(11).Selected = True
        Case &HFFFF&
            cboTags.ComboItems.Item(12).Selected = True
        Case &HFF0000
            cboTags.ComboItems.Item(13).Selected = True
        Case &HFF00FF
            cboTags.ComboItems.Item(14).Selected = True
        Case &HFFFF00
            cboTags.ComboItems.Item(15).Selected = True
        Case &HFFFFFF
            cboTags.ComboItems.Item(16).Selected = True
        Case Else
            cboTags.ComboItems.Item(17).Selected = True
            cboTags.Tag = HTML_Color
    End Select
    
    Select Case Comment_Color
        Case &H0&
            cboComments.ComboItems.Item(1).Selected = True
        Case &H80&
            cboComments.ComboItems.Item(2).Selected = True
        Case &H8000&
            cboComments.ComboItems.Item(3).Selected = True
        Case &H8080&
            cboComments.ComboItems.Item(4).Selected = True
        Case &H800000
            cboComments.ComboItems.Item(5).Selected = True
        Case &H800080
            cboComments.ComboItems.Item(6).Selected = True
        Case &H808000
            cboComments.ComboItems.Item(7).Selected = True
        Case &H808080
            cboComments.ComboItems.Item(8).Selected = True
        Case &HC0C0C0
            cboComments.ComboItems.Item(9).Selected = True
        Case &HFF&
            cboComments.ComboItems.Item(10).Selected = True
        Case &HFF00&
            cboComments.ComboItems.Item(11).Selected = True
        Case &HFFFF&
            cboComments.ComboItems.Item(12).Selected = True
        Case &HFF0000
            cboComments.ComboItems.Item(13).Selected = True
        Case &HFF00FF
           cboComments.ComboItems.Item(14).Selected = True
        Case &HFFFF00
            cboComments.ComboItems.Item(15).Selected = True
        Case &HFFFFFF
            cboComments.ComboItems.Item(16).Selected = True
        Case Else
            cboComments.ComboItems.Item(17).Selected = True
            cboComments.Tag = Comment_Color
    End Select
    
    Select Case ExitScreenParameters.Parameters
        Case NoDontSave
            optDontSave.Value = True
            Label1.Enabled = False
            txtSaveTo.Enabled = False
        Case NoSave
            optSave.Value = True
            Label1.Enabled = False
            txtSaveTo.Enabled = False
        Case NoSaveAll
            optSaveAll.Value = True
            txtSaveTo.Text = ExitScreenParameters.SaveAllLocation
        Case Yes
            ckShowExitScreen.Value = vbChecked
            optDontSave.Enabled = False
            optSave.Enabled = False
            optSaveAll.Enabled = False
            Label1.Enabled = False
            txtSaveTo.Enabled = False
    End Select
    
    cmdApply.Enabled = False
    Changed = False
End Sub

Private Sub optDontSave_Click()
    Label1.Enabled = False
    txtSaveTo.Enabled = False
    
    cmdApply.Enabled = True
End Sub

Private Sub optNewPage_Click()
    txtDocumentName.Text = ""
    txtDocumentName.Enabled = False
    
    cmdApply.Enabled = True
End Sub

Private Sub optOther_Click()
    txtDocumentName.Enabled = True
    
    cmdApply.Enabled = True
End Sub

Private Sub optSave_Click()
    Label1.Enabled = False
    txtSaveTo.Enabled = False
    
    cmdApply.Enabled = True
End Sub

Private Sub optSaveAll_Click()
    Label1.Enabled = True
    txtSaveTo.Enabled = True
    
    cmdApply.Enabled = True
End Sub

Private Sub optUntitled_Click()
    txtDocumentName.Text = ""
    txtDocumentName.Enabled = False
    
    cmdApply.Enabled = True
End Sub

Private Sub TabStrip_Click()
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    
    Select Case TabStrip.SelectedItem.Key
        Case "General"
            Picture3.Visible = True
        Case "Default Document Name"
            Picture4.Visible = True
        Case "Appearance"
            Picture2.Visible = True
        Case "Exit Wizard"
            Picture1.Visible = True
    End Select
End Sub

Private Function GetColorFromString(ByVal sColor As String) As Long
    sColor = StrConv(sColor, vbUpperCase)
    
    Select Case sColor
        Case "BLACK"
            GetColorFromString = &H0&
        Case "MAROON"
            GetColorFromString = &H80&
        Case "GREEN"
            GetColorFromString = &H8000&
        Case "OLIVE"
            GetColorFromString = &H8080&
        Case "NAVY"
            GetColorFromString = &H800000
        Case "PURPLE"
            GetColorFromString = &H800080
        Case "TEAL"
            GetColorFromString = &H808000
        Case "GRAY"
            GetColorFromString = &H808080
        Case "SILVER"
            GetColorFromString = &HC0C0C0
        Case "RED"
            GetColorFromString = &HFF&
        Case "LIME"
            GetColorFromString = &HFF00&
        Case "YELLOW"
            GetColorFromString = &HFFFF&
        Case "BLUE"
            GetColorFromString = &HFF0000
        Case "FUSCHIA"
            GetColorFromString = &HFF00FF
        Case "AQUA"
            GetColorFromString = &HFFFF00
        Case "WHITE"
            GetColorFromString = &HFFFFFF
        Case Else
            GetColorFromString = &H0&
    End Select
End Function

Private Sub txtDocumentName_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtFontSize_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtSaveTo_Change()
    cmdApply.Enabled = True
End Sub
