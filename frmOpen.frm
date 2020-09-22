VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox picRecent 
      BorderStyle     =   0  'None
      Height          =   3180
      Left            =   195
      ScaleHeight     =   3180
      ScaleWidth      =   6690
      TabIndex        =   0
      Top             =   480
      Width           =   6695
      Begin VB.PictureBox picTmp 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   1800
         ScaleHeight     =   360
         ScaleWidth      =   1560
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   1560
      End
      Begin ComctlLib.ListView lRecent 
         Height          =   3135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ilsRecent"
         SmallIcons      =   "ilsRecent"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Folder"
            Object.Width           =   2540
         EndProperty
      End
      Begin ComctlLib.ImageList ilsRecent 
         Left            =   5520
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin VB.PictureBox picExisting 
      BorderStyle     =   0  'None
      Height          =   4020
      Left            =   195
      ScaleHeight     =   4020
      ScaleWidth      =   6690
      TabIndex        =   5
      Top             =   480
      Width           =   6695
      Begin VB.ComboBox Combo 
         Height          =   315
         ItemData        =   "frmOpen.frx":0000
         Left            =   1200
         List            =   "frmOpen.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3240
         Width           =   3975
      End
      Begin ComctlLib.TreeView tExisting 
         Height          =   3135
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5530
         _Version        =   327682
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ilsExisting"
         Appearance      =   1
      End
      Begin ComctlLib.ImageList ilsExisting 
         Left            =   6000
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmOpen.frx":002B
               Key             =   "MyComputer"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmOpen.frx":037D
               Key             =   "Desktop"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label 
         Caption         =   "File Types : "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3280
         Width           =   855
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Existing"
            Key             =   "Existing"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Existing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Recent"
            Key             =   "Recent"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Recent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOpen"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim tmp As String
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim File As File

    If TabStrip.Tabs(1).Selected = True Then GoTo Point_Existing
    If TabStrip.Tabs(2).Selected = True Then GoTo Point_Recent
    
    Exit Sub

Point_Recent:
On Error GoTo Recent_Err
    If lRecent.SelectedItem.Key = "" Then
        MsgBox "Please select a file to open.", vbOKOnly + vbInformation, "Error"
        Exit Sub
    Else
        Set strm = fso.OpenTextFile(lRecent.SelectedItem.SubItems(1) & "\" & lRecent.SelectedItem.Text, ForReading)
On Error GoTo SkipRead
        tmp = strm.ReadAll
        strm.Close
SkipRead:
        Set File = fso.GetFile(lRecent.SelectedItem.SubItems(1) & "\" & lRecent.SelectedItem.Text)
        
        Unload Me
        LoadPage tmp, File.Name, False, File.Path
    End If
    
    Exit Sub
    
Recent_Err:
    MsgBox "The selected file cannot be found.", vbOKOnly + vbInformation, "Error"
    
    Exit Sub

Point_Existing:
    If Mid(tExisting.SelectedItem.Key, 1, 4) = "FILE" Then
        Set strm = fso.OpenTextFile(Mid(tExisting.SelectedItem.Key, 5, Len(tExisting.SelectedItem.Key) - 4), ForReading)
        tmp = strm.ReadAll
        strm.Close
        
        Set File = fso.GetFile(Mid(tExisting.SelectedItem.Key, 5, Len(tExisting.SelectedItem.Key) - 4))
        
        Dim Item As ListItem
        Dim Continue As Boolean
        
        Continue = True
        
        For Each Item In lRecent.ListItems
            If Item.Text = File.Name Then Continue = False
        Next
        
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
        
        Unload Me
        LoadPage tmp, File.Name, False, File.Path
    Else
        MsgBox "Please select a file to open.", vbOKOnly + vbInformation, "Error"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    picExisting.Visible = True
    picRecent.Visible = False

    Combo.Text = Combo.List(0)

    lRecent.ColumnHeaders(2).Width = lRecent.Width - lRecent.ColumnHeaders(1).Width - 680

On Error GoTo err
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim Item As ListItem
    Dim tmp1 As String
    Dim tmp2 As String
    Dim File As File
    Dim Folder As Folder
    Dim Folder2 As Folder
    Dim Drive As Drive
    
    Set strm = fso.OpenTextFile(App.Path & "\Recent", ForReading)
    
On Error Resume Next
    Do Until strm.AtEndOfStream
        tmp1 = strm.ReadLine
        tmp2 = strm.ReadLine
        
        If Right(tmp2, 1) <> "\" Then tmp2 = tmp2 & "\"

        Set Item = lRecent.ListItems.Add(, tmp1, tmp1, ExtractIcon(tmp2 & tmp1, ilsRecent, picTmp, 16), ExtractIcon(tmp2 & tmp1, ilsRecent, picTmp, 16))
        Item.SubItems(1) = tmp2
    Loop
    
    Set Drive = fso.GetDrive("c:\")
    Set Folder2 = fso.GetFolder(Drive.RootFolder)
    
    tExisting.Nodes.Add , , "Desktop", "Desktop", "Desktop", "Desktop"
    tExisting.Nodes.Add "Desktop", tvwChild, "My Computer", "My Computer", "MyComputer", "MyComputer"
    
    For Each Drive In fso.Drives
        If Drive.IsReady = True Then
            If Drive.DriveType = Fixed Then
                tExisting.Nodes.Add "My Computer", tvwChild, "BEFOREDRIVE" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16), ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                
                Set Folder2 = fso.GetFolder(Drive.RootFolder)
                
                For Each Folder In Folder2.SubFolders
                    tExisting.Nodes.Add "BEFOREDRIVE" & Drive.DriveLetter, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsExisting, picTmp, 16), ExtractIcon(Folder.Path, ilsExisting, picTmp, 16)
                Next
                
                For Each File In Folder2.Files
                    If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                        tExisting.Nodes.Add "BEFOREDRIVE" & Drive.DriveLetter, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, ilsExisting, picTmp, 16), ExtractIcon(File.Path, ilsExisting, picTmp, 16)
                    End If
                Next
            Else
                tExisting.Nodes.Add "My Computer", tvwChild, "DRIVE" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16), ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                
                Set Folder2 = fso.GetFolder(Drive.RootFolder)
                
                For Each Folder In Folder2.SubFolders
                    tExisting.Nodes.Add "DRIVE" & Drive.DriveLetter, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsExisting, picTmp, 16), ExtractIcon(Folder.Path, ilsExisting, picTmp, 16)
                Next
                
                For Each File In Folder2.Files
                    If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                        tExisting.Nodes.Add "DRIVE" & Drive.DriveLetter, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, ilsExisting, picTmp, 16), ExtractIcon(File.Path, ilsExisting, picTmp, 16)
                    End If
                Next
            End If
        Else
            tExisting.Nodes.Add "My Computer", tvwChild, "DRIVE" & Drive.DriveLetter, Drive.DriveLetter, ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16), ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
        End If
    Next
    
    tExisting.Nodes.Item(1).Expanded = True
    tExisting.Nodes.Item(2).Expanded = True
    tExisting.Nodes.Item(4).Expanded = True
    
    Set fso = Nothing
    Set strm = Nothing
    Set Item = Nothing
    Set File = Nothing
    Set Folder = Nothing
    Set Folder2 = Nothing
    Set Drive = Nothing
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(App.Path & "\Recent")
    strm.Close
    
    Set File = fso.GetFile(App.Path & "\Recent")
    
    File.Attributes = Hidden + System
    
    Set File = Nothing
    
    Form_Load
End Sub

Private Sub TabStrip_Click()
    picExisting.Visible = False
    picRecent.Visible = False
    
    If TabStrip.Tabs(1).Selected = True Then picExisting.Visible = True
    If TabStrip.Tabs(2).Selected = True Then picRecent.Visible = True
End Sub

Private Sub tExisting_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
    If Mid(Node.Key, 1, 6) = "BEFORE" Then
        Node.Expanded = True
    Else
        tExisting.Visible = False

        Dim Folder As Folder
        Dim File As File
        Dim fso As New FileSystemObject
        Dim Folder2 As Folder
        Dim Drive As Drive
        Dim tNode As Node

        If Mid(Node.Key, 1, 6) = "FOLDER" Then
            Set Folder2 = fso.GetFolder(Mid(Node.Key, 7, Len(Node.Key) - 6))

            For Each Folder In Folder2.SubFolders
                tExisting.Nodes.Add Node.Key, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsExisting, picTmp, 16), ExtractIcon(Folder.Path, ilsExisting, picTmp, 16)
            Next

            For Each File In Folder2.Files
                If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                    tExisting.Nodes.Add Node.Key, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, ilsExisting, picTmp, 16), ExtractIcon(File.Path, ilsExisting, picTmp, 16)
                End If
            Next

            Node.Key = "BEFORE" & Node.Key
            Node.Expanded = True
        Else
            If Mid(Node.Key, 1, 5) = "DRIVE" Then
                Set Drive = fso.GetDrive(Mid(Node.Key, 6, Len(Node.Key) - 5))
                
                Do Until Node.Children = 0
                    tExisting.Nodes.Remove Node.Child.Index
                Loop
                
                If Drive.IsReady = True Then
                    Set Folder2 = Drive.RootFolder

                    For Each Folder In Folder2.SubFolders
                        tExisting.Nodes.Add Node.Key, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsExisting, picTmp, 16), ExtractIcon(Folder.Path, ilsExisting, picTmp, 16)
                    Next

                    For Each File In Folder2.Files
                        If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                            tExisting.Nodes.Add Node.Key, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, ilsExisting, picTmp, 16), ExtractIcon(File.Path, ilsExisting, picTmp, 16)
                        End If
                    Next
                    
                    Node.Text = Drive.VolumeName & " (" & Drive.DriveLetter & ")"
                    Node.SelectedImage = ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                    Node.Image = ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                    Node.Expanded = True
                Else
                    Node.Text = Drive.DriveLetter
                    Node.SelectedImage = ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                    Node.Image = ExtractIcon(Drive.Path & "\", ilsExisting, picTmp, 16)
                    
                    MsgBox "Drive is not ready.", vbOKOnly + vbInformation, "Error"
                End If
            End If
        End If

        Set Folder = Nothing
        Set File = Nothing
        Set fso = Nothing
        Set Folder2 = Nothing

        tExisting.Visible = True
    End If
End Sub
