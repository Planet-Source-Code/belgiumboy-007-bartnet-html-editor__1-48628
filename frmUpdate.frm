VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check For Update"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   840
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   360
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Min             =   1e-4
   End
   Begin VB.CommandButton cmdCheckForUpdate 
      Caption         =   "Check For Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmUpdate"
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

Private Type FT
    Version As String
    Url As String
    FileSize As String
    FileName As String
    FileLocation As String
End Type

Private myData() As Byte

Private CurrStatus As Integer
Private tmp As String
Private FileDetails As FT

Private fso As New FileSystemObject
Private strm As TextStream

Private Sub cmdCancel_Click()
    ProgressBar.Value = 0
    Unload Me
End Sub

Private Sub cmdCheckForUpdate_Click()
    cmdCheckForUpdate.Enabled = False
    Timer.Enabled = True
End Sub

Private Sub Form_Load()
    ProgressBar.Min = 0
    ProgressBar.Max = 10
    ProgressBar.Value = 0
    Label.Caption = "Click the Check For Update button to start the process."
    CurrStatus = 0
    Timer2.Enabled = True
End Sub

Private Sub Timer_Timer()
On Local Error GoTo err
    Select Case CurrStatus
        Case 0
            Label.Caption = "Connecting to www.bartnet.be."
        Case 1
            myData() = Inet.OpenURL("http://www.bartnet.be/765NU189I7CNE756YB651MJDLSNFDYSU.BartNet", icByteArray)
        Case 3
            Label.Caption = "Processing data."
        Case 4
            Open fso.GetSpecialFolder(TemporaryFolder) & "\765NU189I7CNE756YB651MJDLSNFDYSU.BartNet" For Binary Access Write As #1
            Put #1, , myData()
            Close #1
        Case 5
            Set strm = fso.OpenTextFile(fso.GetSpecialFolder(TemporaryFolder) & "\765NU189I7CNE756YB651MJDLSNFDYSU.BartNet", ForReading)
            
            Do Until strm.AtEndOfStream
                tmp = strm.ReadLine
                
                If Mid(tmp, 1, 10) = "NEWVERSION" Then
                    FileDetails.Version = Mid(tmp, 14, Len(tmp) - 13)
                Else
                    If Mid(tmp, 1, 3) = "URL" Then
                        FileDetails.Url = Mid(tmp, 7, Len(tmp) - 6)
                    Else
                        If Mid(tmp, 1, 8) = "FILESIZE" Then
                            FileDetails.FileSize = Mid(tmp, 12, Len(tmp) - 11)
                        Else
                            If Mid(tmp, 1, 8) = "FILENAME" Then
                                FileDetails.FileName = Mid(tmp, 12, Len(tmp) - 11)
                            Else
                                If Mid(tmp, 1, 12) = "FILELOCATION" Then
                                    FileDetails.FileLocation = Mid(tmp, 16, Len(tmp) - 15)
                                End If
                            End If
                        End If
                    End If
                End If
            Loop
            
            strm.Close
        Case 6
            If FileDetails.Version = "" Then FileDetails.Version = ProgVersion
            
            If ProgVersion = FileDetails.Version Then
                MsgBox "You already have the latest version of BartNet HTML Editor.", vbOKOnly + vbInformation, "Result"
                cmdCancel_Click
            Else
                If MsgBox("There is a new version of BartNet HTML Editor available, details follow :" & vbCrLf & vbCrLf & "Version : " & FileDetails.Version & vbCrLf & "URL : " & FileDetails.Url & vbCrLf & vbCrLf & "File Name : " & FileDetails.FileName & vbCrLf & "File Size : " & FileDetails.FileSize & vbCrLf & vbCrLf & "BartNet HTML Editor can download the new version for you, would you like it to download the new version for you ?", vbYesNo + vbQuestion, "Result") = vbNo Then
                    cmdCancel_Click
                End If
            End If
        Case 7
            Label.Caption = "Downloading new version."
        Case 8
            myData() = Inet.OpenURL(FileDetails.FileLocation, icByteArray)
        Case 9
            Open App.Path & "\BartNet HTML Editor " & FileDetails.Version & ".zip" For Binary Access Write As #1
            Put #1, , myData
            Close #1
        Case 10
            If MsgBox("Download successful." & vbCrLf & vbCrLf & "Open downloaded WinZip file now ?", vbYesNo + vbQuestion, "Result") = vbYes Then
                frmMDI.WebBrowser.Navigate App.Path & "\BartNet HTML Editor " & FileDetails.Version & ".zip", , "_new"
                cmdCancel_Click
            Else
                cmdCancel_Click
            End If
            
            Timer.Enabled = False
    End Select
    
    If CurrStatus < 10 Then
        CurrStatus = CurrStatus + 1
        ProgressBar.Value = CurrStatus
    End If
    
    Exit Sub
    
err:
    MsgBox "An error has occured during the process.  Please try again later.", vbOKOnly + vbInformation, "Error"
    cmdCancel_Click
End Sub

Private Sub Timer2_Timer()
    If ProgressBar.Value <> 0 Then ProgressBar.Value = 0
    Timer2.Enabled = False
End Sub
