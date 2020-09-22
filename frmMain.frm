VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   6540
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   3360
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   6135
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   975
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
         ExtentX         =   1720
         ExtentY         =   1720
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
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
   End
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin RichTextLib.RichTextBox RichTextBox 
         Height          =   1335
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2355
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":038A
      End
   End
   Begin RichTextLib.RichTextBox rtbTemp 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":040C
   End
End
Attribute VB_Name = "frmMain"
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

Public Saved As Boolean
Public SavedBefore As Boolean
Public SavedBeforePath As String

Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

Public FirstLoad As Boolean

Private Sub Form_GotFocus()
    frmMDI.UpdateView
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    frmMDI.UpdateView
End Sub

Private Sub Form_LostFocus()
    frmMDI.UpdateView
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMDI.UpdateView
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Timer.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        picPreview.Width = Me.ScaleWidth
        picEdit.Width = picPreview.Width
        WebBrowser.Width = picPreview.ScaleWidth
        RichTextBox.Width = picEdit.ScaleWidth
        
        picPreview.Height = Me.ScaleHeight
        picEdit.Height = picPreview.Height
        WebBrowser.Height = picPreview.ScaleHeight
        RichTextBox.Height = picEdit.ScaleHeight
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormsCount = FormsCount - 1
End Sub

Public Sub RichTextBox_Change()
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = RichTextBox.TextRTF
    End If
    
    Saved = False
    
    If DisableTagColoring Then Exit Sub
        
    With RichTextBox
        SelCursor = .SelStart
        SelLength = .SelLength
            
        rtbTemp = RichTextBox
            
        If FirstLoad Then ColorTags rtbTemp, 0 Else ColorTags rtbTemp, .SelStart
        
    On Error Resume Next
        RichTextBox = rtbTemp
        .SetFocus
        
        .SelStart = SelCursor
        .SelLength = SelLength
    End With
End Sub

Private Sub RichTextBox_GotFocus()
    frmMDI.UpdateView
End Sub

Private Sub RichTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RichTextBox.SelFontName = DefaultFontType
    RichTextBox.SelFontSize = DefaultFontSize
End Sub

Private Sub RichTextBox_LostFocus()
    frmMDI.UpdateView
End Sub

Private Sub RichTextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RichTextBox.SelFontName = DefaultFontType
    RichTextBox.SelFontSize = DefaultFontSize
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    frmMDI.mnuClose_Click
End Sub

Public Sub Undo()
    If gintIndex = 0 Then Exit Sub
    
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
On Error Resume Next
    RichTextBox.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Public Sub Redo()
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
On Error Resume Next
    RichTextBox.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub WebBrowser_GotFocus()
    frmMDI.UpdateView
End Sub

Private Sub WebBrowser_LostFocus()
    frmMDI.UpdateView
End Sub

Private Sub WebBrowser_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    If picPreview.Visible = False Then Exit Sub
    
    If frmMDI.ActiveForm.Name = Me.Name Then
        frmMDI.txtStatus.Text = "Done"
        frmMDI.pStatus.Visible = False
    End If
End Sub

Private Sub WebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If picPreview.Visible = False Then Exit Sub

    If Progress = 0 Then Progress = 1
    
    If frmMDI.ActiveForm.Name = Me.Name Then
        frmMDI.pStatus.Min = 0
        frmMDI.pStatus.Max = ProgressMax
        frmMDI.pStatus.Value = Progress
    
        If Progress <> 1 And frmMDI.StatusBar.Visible = True Then frmMDI.pStatus.Visible = True Else frmMDI.pStatus.Visible = False
    End If
End Sub

Private Sub WebBrowser_StatusTextChange(ByVal Text As String)
    If picPreview.Visible = False Then Exit Sub
    
    If frmMDI.ActiveForm.Name = Me.Name Then
        frmMDI.txtStatus.Text = Text
    End If
End Sub
