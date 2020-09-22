VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.CheckBox Check 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Find What :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
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

Private FirstTime As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    Dim Found As Integer
    
    If FirstTime = True Then
        frmMDI.ActiveForm.RichTextBox.SelStart = 0
        FirstTime = False
    End If
    
    If Check.Value = vbChecked Then
        Found = frmMDI.ActiveForm.RichTextBox.Find(Text.Text, frmMDI.ActiveForm.RichTextBox.SelStart + Len(Text.Text), , rtfMatchCase)
    Else
        Found = frmMDI.ActiveForm.RichTextBox.Find(Text.Text, frmMDI.ActiveForm.RichTextBox.SelStart + Len(Text.Text))
    End If
    
    If Found = True Then
        MsgBox "Search text not found.", vbOKOnly + vbInformation + vbSystemModal, "Not Found"
    Else
        frmMDI.ActiveForm.RichTextBox.SetFocus
    End If
End Sub

Private Sub Form_Load()
    FormOnTop Me.hWnd, True
    FirstTime = True
End Sub
