VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Select Location"
      Height          =   375
      Left            =   2340
      TabIndex        =   16
      Top             =   24000
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Progress"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   45600
      Width           =   6375
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label lblProgress 
         Caption         =   "Label4"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5040
      Top             =   32400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ListBox lstCreatedIndex 
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   55200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstOpenedIndex 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   55200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Created"
      Height          =   4575
      Left            =   3360
      TabIndex        =   1
      Top             =   84000
      Width           =   3135
      Begin VB.ListBox lstCreated 
         Height          =   3435
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "These documents have not been saved before, you will be prompted for a location to save each checked document into."
         Height          =   800
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opened"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   84000
      Width           =   3135
      Begin VB.ListBox lstOpened 
         Height          =   3435
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "These documents have been saved before."
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   2895
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Done, click Finish to exit."
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblLocation 
      Caption         =   "Label4"
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   12000
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   $"frmExit.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   12000
      Width           =   6375
   End
End
Attribute VB_Name = "frmExit"
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

Private CurrentDocument As Integer
Private TotalDocuments As Integer
Private Continue As Boolean
Private Currenti As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdContinue_Click()
    If cmdContinue.Caption = "Finish" Then
        cmdExit_Click
    Else
        SetStatus GetStatus + 1
    End If
End Sub

Private Sub cmdExit_Click()
    Unload frmMDI
End Sub

Private Sub cmdOK_Click()
On Error GoTo CommErr
    CommonDialog.ShowSave
On Error Resume Next
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim tmp As String
    Dim tmpForm As Form
    
    Set tmpForm = Forms(lstCreatedIndex.List(Currenti))
    tmp = tmpForm.RichTextBox.Text
    
    Set strm = fso.CreateTextFile(CommonDialog.FileName, True)
    strm.Write tmp
    strm.Close
    
    ProgressBar.Value = ProgressBar.Value + 1
    lblProgress.Caption = CurrentDocument & " of " & TotalDocuments & " files saved."
    
    Continue = True
    
    Exit Sub
    
CommErr:
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Me.Height = 6480
    Me.Width = 6690
    
    Label1.Top = 120
    Frame1.Top = 840
    Frame2.Top = 840
    
    Frame3.Top = 4560
    cmdOK.Top = 2400
    lblLocation.Top = 120
    
    Label4.Top = 120

    Dim i As Integer
    Dim tmpForm As Form
    Dim Continue As Boolean
    
    Continue = False
    
    Do Until i = FormsCount + 2
        If TypeOf Forms(i) Is frmMain Then
            If Forms(i).Saved = False Then
                Set tmpForm = Forms(i)
                
                If tmpForm.SavedBefore = True Then
                    lstOpened.AddItem tmpForm.Caption
                    lstOpenedIndex.AddItem i
                Else
                    lstCreated.AddItem tmpForm.Caption
                    lstCreatedIndex.AddItem i
                End If
                
                Continue = True
            End If
        End If
        
        i = i + 1
    Loop
    
    If Continue = True Then
        If ExitScreenParameters.Parameters = Yes Then
            SetStatus 1
        Else
            If ExitScreenParameters.Parameters = NoDontSave Then
                cmdExit_Click
            Else
                Dim fso As New FileSystemObject
                Dim strm As TextStream
                
                i = 0
                
                If ExitScreenParameters.Parameters = NoSave Then
                    Do Until i = lstOpened.ListCount
                        Set tmpForm = Forms(lstOpenedIndex.List(i))
                        tmp = tmpForm.RichTextBox.Text
                        
                        Set strm = fso.CreateTextFile(tmpForm.SavedBeforePath, True)
                        strm.Write tmp
                        strm.Close
                        
                        i = i + 1
                    Loop
                    
                    cmdExit_Click
                Else
                    Do Until i = lstOpened.ListCount
                        Set tmpForm = Forms(lstOpenedIndex.List(i))
                        tmp = tmpForm.RichTextBox.Text
                        
                        Set strm = fso.CreateTextFile(tmpForm.SavedBeforePath, True)
                        strm.Write tmp
                        strm.Close
                        
                        i = i + 1
                    Loop
                    
                    i = 0
                    
                    Do Until i = lstCreated.ListCount
                        Set tmpForm = Forms(lstCreatedIndex.List(i))
                        tmp = tmpForm.RichTextBox.Text
                        
                        Set strm = fso.CreateTextFile(ExitScreenParameters.SaveAllLocation & tmpForm.Caption & ".html", True)
                        strm.Write tmp
                        strm.Close
                        
                        i = i + 1
                    Loop
                    
                    cmdExit_Click
                End If
            End If
        End If
    Else
        cmdExit_Click
    End If
End Sub

Private Sub SetStatus(ByVal intStat As Integer)
    Label1.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    lblLocation.Visible = False
    cmdOK.Visible = False
    Label4.Visible = False

    Select Case intStat
        Case 1
            Label1.Visible = True
            Frame1.Visible = True
            Frame2.Visible = True
            
            Exit Sub
        Case 2
            Dim intAmount As Integer
            Dim a As Integer
            
            Do Until a = lstCreated.ListCount
                If lstCreated.Selected(a) = True Then intAmount = intAmount + 1
                
                a = a + 1
            Loop
            
            If intAmount <> 0 Then
                MsgBox "You will now be prompted for a location for every one of the " & intAmount & " documents you checked and haven't been saved yet.", vbOKOnly + vbInformation, "Info"
            End If
            
            SetStatus 3
            
            Exit Sub
        Case 3
            Frame3.Visible = True
            lblLocation.Visible = True
            cmdOK.Visible = True
            cmdContinue.Enabled = False
            
            Dim i As Integer
            Dim intAmount1 As Integer
            Dim intAmount2 As Integer
            
            Do Until i = lstCreated.ListCount
                If lstCreated.Selected(i) = True Then intAmount1 = intAmount1 + 1
                
                i = i + 1
            Loop
            
            i = 0
            
            Do Until i = lstOpened.ListCount
                If lstOpened.Selected(i) = True Then intAmount2 = intAmount2 + 1
                
                i = i + 1
            Loop
            
            If intAmount1 + intAmount2 = 0 Then
                SetStatus 4
            Else
                ProgressBar.Min = 0
                ProgressBar.Max = intAmount1 + intAmount2
                ProgressBar.Value = 0
                lblProgress.Caption = "0 of " & intAmount1 + intAmount2 & " files saved."
                TotalDocuments = intAmount1 + intAmount2
            
                i = 0
            
                CommonDialog.Filter = "Internet Documents (HTML)|*.htm;*.html"
    
                Do Until i = lstCreated.ListCount
                    If lstCreated.Selected(i) = True Then
                        CommonDialog.FileName = lstCreated.List(i)
                        lblLocation.Caption = "The next file is " & lstCreated.List(i) & "." & vbCrLf & vbCrLf & "Please select a location to save this file to by pressing the Select Location button below."
                        Currenti = i
                        CurrentDocument = CurrentDocument + 1
                        
                        Continue = False
                        
                        Do Until Continue = True
                            DoEvents
                        Loop
                    End If
                
                    i = i + 1
                Loop
                
                cmdOK.Enabled = False
                
                If CurrentDocument = TotalDocuments Then
                    SetStatus 4
                Else
                    i = 0
                    
                    Dim fso As New FileSystemObject
                    Dim strm As TextStream
                    Dim tmp As String
                    Dim tmpForm As Form
                        
                    Do Until i = lstOpened.ListCount
                        If lstOpened.Selected(i) = True Then
                            lblLocation.Caption = "The next file is " & lstOpened.List(i) & "." & vbCrLf & vbCrLf & "Please wait."
                        
                            Set tmpForm = Forms(lstOpenedIndex.List(i))
                            tmp = tmpForm.RichTextBox.Text
                        
                            Set strm = fso.CreateTextFile(tmpForm.SavedBeforePath, True)
                            strm.Write tmp
                            strm.Close
                            
                            CurrentDocument = CurrentDocument + 1
                            ProgressBar.Value = ProgressBar.Value + 1
                            lblProgress.Caption = CurrentDocument & " of " & TotalDocuments & " files saved."
                        End If
                        
                        i = i + 1
                    Loop
                    
                    SetStatus 4
                End If
            End If
        Case 4
            cmdContinue.Enabled = True
            cmdContinue.Caption = "Finish"
            cmdExit.Enabled = False
            Label4.Visible = True
    End Select
    
    Exit Sub
    

End Sub

Private Function GetStatus()
    If Label1.Visible = True And Frame1.Visible = True And Frame2.Visible = True Then
        GetStatus = 1
        Exit Function
    End If
End Function
