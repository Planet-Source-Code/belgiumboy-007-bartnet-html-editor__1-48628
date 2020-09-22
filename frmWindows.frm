VERSION 5.00
Begin VB.Form frmWindows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GoTo Window"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstIndex 
      Height          =   450
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox List 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmWindows"
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

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'On Error Resume Next
    Dim tmpForm As Form
    Dim i As Integer
    Dim iToUse As Integer
    
    iToUse = 1
    
    Do Until i = List.ListCount
        If List.List(i) = List.Text Then
            iToUse = lstIndex.List(i)
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    
    Set tmpForm = Forms(iToUse)
    
    Me.Hide
    
    tmpForm.SetFocus
    tmpForm.RichTextBox.SetFocus
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim tmpForm As Form
    Dim Continue As Boolean
    
    Continue = False
    
    Do Until i = FormsCount + 2
        If TypeOf Forms(i) Is frmMain Then
            Set tmpForm = Forms(i)
            
            List.AddItem tmpForm.Caption
            lstIndex.AddItem i
            
            Continue = True
        End If
        
        i = i + 1
    Loop
    
    If Continue = False Then
        MsgBox "There are no open windows to go to.", vbOKOnly + vbInformation, "Error"
        Unload Me
    End If
End Sub
