VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3315
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   960
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_Load()
    Timer.Enabled = True
End Sub

Private Sub Timer_Timer()
On Error GoTo err
    Dim fso As New FileSystemObject
    Dim strm As TextStream
        
    Set strm = fso.OpenTextFile(App.Path & "\Settings", ForReading)
    
    StartUpScreenShow = strm.ReadLine
    ShowToolBar = strm.ReadLine
    ShowStatusBar = strm.ReadLine
    DocumentName = strm.ReadLine
    DefaultFontType = strm.ReadLine
    DefaultFontSize = strm.ReadLine
    HTML_Color = strm.ReadLine
    Comment_Color = strm.ReadLine
    DisableTagColoring = strm.ReadLine
    DefaultBackGroundColor = strm.ReadLine
    ExitScreenParameters.Parameters = strm.ReadLine
    ExitScreenParameters.SaveAllLocation = strm.ReadLine

    strm.Close
    
    If ShowToolBar = True Then
        frmMDI.Toolbar.Visible = True
        Set frmMDI.img(22).Picture = frmMDI.imgCheck.Picture
    Else
        frmMDI.Toolbar.Visible = False
        Set frmMDI.img(22).Picture = frmMDI.imgUncheck.Picture
    End If
    
    If ShowStatusBar = True Then
        frmMDI.picStatus.Visible = True
        Set frmMDI.img(24).Picture = frmMDI.imgCheck.Picture
    Else
        frmMDI.picStatus.Visible = False
        Set frmMDI.img(24).Picture = frmMDI.imgUncheck.Picture
    End If
    
    FirstLoad = False
    
    Me.Hide
    
    If StartUpScreenShow = True Then
        If LoadValues = True Then frmStartUp.Show vbModal, frmMDI
    Else
        PageNumber = PageNumber + 1
        
        Dim tmp As String
        
        tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        
        LoadPage tmp, DocumentName & " " & PageNumber
    End If
    
    Unload Me
    
    Timer.Enabled = False
    
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
    
    Timer_Timer
End Sub

