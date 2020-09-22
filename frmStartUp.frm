VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Document"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check 
      Caption         =   "Show this screen at startup."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.PictureBox picRecent 
      BorderStyle     =   0  'None
      Height          =   3180
      Left            =   193
      ScaleHeight     =   3180
      ScaleWidth      =   6690
      TabIndex        =   10
      Top             =   480
      Width           =   6695
      Begin VB.PictureBox picTmp 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   1800
         ScaleHeight     =   360
         ScaleWidth      =   1560
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1560
      End
      Begin ComctlLib.ListView lRecent 
         Height          =   3135
         Left            =   0
         TabIndex        =   11
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5753
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   5753
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox picNew 
      BorderStyle     =   0  'None
      Height          =   3180
      Left            =   193
      ScaleHeight     =   3180
      ScaleWidth      =   6690
      TabIndex        =   2
      Top             =   480
      Width           =   6695
      Begin ComctlLib.ListView lNew 
         Height          =   3135
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5530
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         Icons           =   "ilsNew"
         SmallIcons      =   "ilsNew"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ImageList ilsNew 
         Left            =   6000
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmStartUp.frx":0000
               Key             =   "Frames"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmStartUp.frx":0C52
               Key             =   "Normal"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picExisting 
      BorderStyle     =   0  'None
      Height          =   4020
      Left            =   193
      ScaleHeight     =   4020
      ScaleWidth      =   6690
      TabIndex        =   6
      Top             =   480
      Width           =   6695
      Begin VB.ComboBox Combo 
         Height          =   315
         ItemData        =   "frmStartUp.frx":18A4
         Left            =   1200
         List            =   "frmStartUp.frx":18AB
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
      Begin VB.Label Label 
         Caption         =   "File Types : "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3280
         Width           =   855
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
               Picture         =   "frmStartUp.frx":18CF
               Key             =   "MyComputer"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmStartUp.frx":1C21
               Key             =   "Desktop"
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   4575
      Left            =   113
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Key             =   "New"
            Object.Tag             =   ""
            Object.ToolTipText     =   "New"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Existing"
            Key             =   "Existing"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Existing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Recent"
            Key             =   "Recent"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Recent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStartUp"
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

On Error GoTo err
    If Check.Value <> vbChecked Then
        Set strm = fso.OpenTextFile(App.Path & "\Settings", ForWriting)
        
        strm.WriteLine "False"
        strm.WriteLine ShowToolBar
        strm.WriteLine ShowStatusBar
        strm.WriteLine DocumentName
        strm.WriteLine DefaultFontType
        strm.WriteLine DefaultFontSize
        strm.WriteLine HTML_Color
        strm.WriteLine Comment_Color
        strm.WriteLine DisableTagColoring
        strm.WriteLine DefaultBackGroundColor
        strm.WriteLine ExitScreenParameters.Parameters
        strm.WriteLine ExitScreenParameters.SaveAllLocation
        
        strm.Close
        
        StartUpScreenShow = False
    End If
    
    If TabStrip.Tabs(1).Selected = True Then GoTo Point_New
    If TabStrip.Tabs(2).Selected = True Then GoTo Point_Existing
    If TabStrip.Tabs(3).Selected = True Then GoTo Point_Recent
    
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
    
    Exit Sub

Point_New:
    PageNumber = PageNumber + 1
    
    tmp = "<html>" & _
        vbCrLf & _
        vbCrLf & "<head>" & _
        vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & _
        vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & _
        "</head>" & _
        vbCrLf & _
        vbCrLf & "<body>"
    
    Select Case lNew.SelectedItem.Key
        Case "Normal Page"
            tmp = tmp & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Bibliography"
            tmp = tmp & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<p><a name=""ALastName"">ALastName</a>, FirstInitial. Year. <em>Title of publication.</em>City, State: Publisher.</p>" & vbCrLf & vbCrLf & "<p><a name=""BLastName"">BLastName</a>, FirstInitial. Year. <em>Title of publication.</em>City, State: Publisher.</p>" & vbCrLf & vbCrLf & "<p><a name=""CLastName"">CLastName</a>, FirstInitial. Year. <em>Title of publication.</em>City, State: Publisher.</p>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Revised: July 04, 2003.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Centered Body"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""3"" width=""100%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""25%""></td>" & vbCrLf & "  <td valign=""top"" width=""50%""><p align=""center""><font size=""5"">Your Heading Goes Here</font></p>" & vbCrLf & "   <p align=""center""><font size=""3""><em>Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </em></font></p>" & _
                vbCrLf & "   <p align=""center""><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font></p>" & _
                vbCrLf & "   <p align=""center""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></p>" & _
                vbCrLf & "   <p align=""center""><font size=""3"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""25%""></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Confirmation Form"
            tmp = tmp & vbCrLf & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<p>Dear [Username],</p>" & vbCrLf & vbCrLf & "<p>Thank you for sending us your [MessageType] about our [Subject]. If you have asked us to contact you, we will be using the following information:</p>" & vbCrLf & _
                vbCrLf & "<blockquote>" & vbCrLf & " <p><strong>E-mail:</strong> [UserEmail]<br>" & vbCrLf & " <strong>Telephone:</strong> [UserTel]<br>" & vbCrLf & " <strong>FAX:</strong> [UserFAX]</p>" & vbCrLf & "</blockquote>" & vbCrLf & vbCrLf & "<p>If any of this information is incorrect, please go back to the feedback form and change it. We thank you for taking the time to help us be a better company.</p>" & _
                vbCrLf & vbCrLf & "<blockquote>" & vbCrLf & " <blockquote>" & vbCrLf & "  <blockquote>" & vbCrLf & "   <p>Sincerely,</p>" & vbCrLf & "   <p><em>Manager, Customer Services</em></p>" & vbCrLf & "  </blockquote>" & vbCrLf & " </blockquote>" & vbCrLf & "</blockquote>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>You may return to the feedback form by using the <em>Back</em> button in your browser.</h5>" & vbCrLf & vbCrLf & "<h5>Revised: July 04, 2003.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Feedback Form"
            tmp = tmp & vbCrLf & vbCrLf & vbCrLf & "<form action=""SomeScript.php"" method=""POST"">" & vbCrLf & " <p><strong>What kind of comment would you like to send?</strong>" & vbCrLf & "  <dl>" & vbCrLf & "   <dd><input type=""radio"" name=""MessageType"" value=""Complaint"">Complaint <input type=""radio"" name=""MessageType"" value=""Problem"">Problem <input type=""radio"" checked name=""MessageType"" value=""Suggestion"">Suggestion <input type=""radio"" name=""MessageType"" value=""Praise"">Praise</dd>" & vbCrLf & "  </dl>" & vbCrLf & " <p><strong>What about us do you want to comment on?</strong>" & vbCrLf & "  <dl>" & vbCrLf & "   <dd><select name=""Subject"" size=""1"">" & vbCrLf & _
                "     <option selected>Web Site</option>" & vbCrLf & "     <option>Company</option>" & vbCrLf & "     <option>Products</option>" & vbCrLf & "     <option>Store</option>" & vbCrLf & "     <option>Employee</option>" & vbCrLf & "     <option>(Other)</option>" & vbCrLf & "    </select> Other: <input type=""text"" size=""26"" maxlength=""256"" name=""SubjectOther"">" & vbCrLf & "   </dd>" & _
                vbCrLf & "  </dl>" & vbCrLf & " <p><strong>Enter your comments in the space provided below:</strong>" & vbCrLf & "  <dl>" & vbCrLf & "   <dd><textarea name=""Comments"" rows=""5"" cols=""42""></textarea></dd>" & vbCrLf & "  </dl>" & vbCrLf & " <p><strong>Tell us how to get in touch with you:</strong>" & vbCrLf & "  <dl>" & vbCrLf & "   <dd><pre>Name     <input type=""text"" size=""35"" maxlength=""256"" name=""Username""> & vbcrlf & ""E-mail   <input type=""text"" size=""35"" maxlength=""256"" name=""UserEmail"">" & vbCrLf & "Tel      <input type=""text"" size=""35"" maxlength=""256"" name=""UserTel"">" & vbCrLf & _
                "FAX      <input type=""text"" size=""35"" maxlength=""256"" name=""UserFAX""> </pre>" & vbCrLf & "   </dd>" & vbCrLf & "  </dl>" & vbCrLf & "  <dl>" & vbCrLf & "   <dd><input type=""checkbox"" name=""ContactRequested"" value=""ContactRequested""> Please contact me as soon as possible regarding this matter.</dd>" & vbCrLf & "  </dl>" & vbCrLf & " <p><input type=""submit"" value=""Submit Comments""> <input type=""reset"" value=""Clear Form""></p>" & vbCrLf & "</form>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © 1997 [OrganizationName]. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Frequently Asked Questions"
            tmp = tmp & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h2><a name=""top"">Table of Contents</a></h2>" & vbCrLf & vbCrLf & "<ol>" & vbCrLf & "  <li><a href=""#how""><strong>How do I ... ?</strong></a></li>" & vbCrLf & "  <li><a href=""#where""><strong>Where can I find ... ?</strong></a></li>" & vbCrLf & "  <li><a href=""#why""><strong>Why doesn't ... ?</strong></a></li>" & vbCrLf & "  <li><a href=""#who""><strong>Who is ... ?</strong></a></li>" & vbCrLf & "  <li><a href=""#what""><strong>What is ... ?</strong></a></li>" & vbCrLf & "  <li><a href=""#when""><strong>When is ... ?</strong></a></li>" & vbCrLf & "</ol>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""how"">How do I ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""where"">Where can I find ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""why"">Why doesn't ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""who"">Who is ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""what"">What is ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h3><a name=""when"">When is ... ?</a></h3>" & vbCrLf & vbCrLf & "<p>[This is the answer to the question.]</p>" & vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © [OrganizationName]. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Guest Book"
            tmp = tmp & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<p>We'd like to know what you think about our web site. Please leave your comments in this public guest book so we can share your thoughts with other visitors.</p>" & _
                vbCrLf & vbCrLf & "<form action=""SomeScript.php"" method=""POST"">" & vbCrLf & "<h2><strong>Add Your Comments</strong></h2>" & vbCrLf & "<p><textarea name=""Comments"" rows=""8"" cols=""52""></textarea></p>" & vbCrLf & "<p><input type=""submit"" value=""Submit Comments""> <input type=""reset"" value=""Clear Comments""><br>" & vbCrLf & "<br>" & vbCrLf & "<em>After you submit your comments, you will need to reload this page with your browser in order to see your additions to the log.</em></p>" & vbCrLf & "</form>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © 1997 by [OrganizationName]. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003 11:17:20.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Narrow, Left-aligned Body"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""15%""></td>" & _
                vbCrLf & "  <td align=""right"" valign=""top"" width=""35%""><font size=""3"" face=""Arial""><strong>Your Heading Goes Here</strong></font>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi.</font>" & vbCrLf & "  </td>" & vbCrLf & "  <td align=""center"" width=""50%""><img src=""" & App.Path & "/earth.jpg"" width=""224"" height=""217""><br>" & _
                vbCrLf & "   <font size=""1""><em><strong>Earth Photo Caption</strong></em></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & vbCrLf & "<p>&nbsp;</p>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Narrow, Right-aligned Body"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""3"" width=""100%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""50%""><img src=""" & App.Path & "/Firework.jpg"" width=""290"" height=""473""><br>" & vbCrLf & "   <font size=""2""><em>fireworks photo caption</em></font>" & vbCrLf & "  </td>" & vbCrLf & "  <td valign=""top"" width=""40%""><font size=""3"" face=""Arial""><strong>Your Heading Goes Here</strong></font>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Contents and Sidebar"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""6"" cellspacing=""0"" width=""100%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td><font size=""6"" face=""Arial"">Your Heading Goes Here</font></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" width=""16%""><font size=""2"" face=""Arial""><strong>SECTION 1</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">Title Goes Here 1<br>" & _
                vbCrLf & "   Title Goes Here 2<br>" & vbCrLf & "   Title Goes Here 3<br>" & vbCrLf & "   Title Goes Here 4</font>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 2</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">Title Goes Here 1<br>" & vbCrLf & "   Title Goes Here 2<br>" & vbCrLf & "   Title Goes Here 3</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 3</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">Title Goes Here 1<br>" & vbCrLf & "   Title Goes Here 2<br>" & vbCrLf & "   Title Goes Here 3<br>" & _
                vbCrLf & "   Title Goes Here 4<br>" & vbCrLf & "   Title Goes Here 5</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 4<br>" & vbCrLf & "   </strong>Title Goes Here 1<br>" & vbCrLf & "   Title Goes Here 2<br>" & vbCrLf & "   Title Goes Here 3<br>" & vbCrLf & "   Title Goes Here 4<br>" & vbCrLf & "   Title Goes Here 5</font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">Title Goes Here 6</font><br>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""8""></td>" & vbCrLf & "  <td valign=""top"" width=""54%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et " & _
                "iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""8""></td>" & vbCrLf & "  <td valign=""top"" width=""18%"" bgcolor=""#000000""><p align=""center""><img src=""" & App.Path & "/Fish.jpg"" width=""147"" height=""81""><br>" & vbCrLf & "   <em><font color=""#FFFFFF"" size=""2"" face=""Arial""><strong>Title Goes Here<br></strong></font>" & _
                vbCrLf & "   <font color=""#FFFFFF"" size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion<br></font></em></p>" & vbCrLf & "   <p align=""center""><em><font color=""#FFFFFF"" size=""2"" face=""Arial""><strong><img src=""" & App.Path & "/Leopard.jpg"" width=""129"" height=""89""><br>" & _
                vbCrLf & "   Title Goes Here<br></strong></font><font color=""#FFFFFF"" size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion<br></font></em></p>" & vbCrLf & "   <p align=""center""><em><font color=""#FFFFFF"" size=""3""><img src=""" & App.Path & "/earth2.jpg"" width=""88"" height=""85""><br></font>" & _
                vbCrLf & "   <font color=""#FFFFFF"" size=""2"" face=""Arial""><strong>Title Goes Here<br></strong></font>" & vbCrLf & "   <font color=""#FFFFFF"" size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion</font></em>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Contents on Left"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""8"" width=""98%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""20%"">&nbsp; </td>" & vbCrLf & "  <td width=""15""></td>" & _
                vbCrLf & "  <td valign=""bottom"" width=""80%""><font size=""5"" face=""Arial"">Your Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" width=""20%""><font size=""2"" face=""Arial""><strong>SECTION 1</strong></font><font size=""3"" face=""Arial""><br>" & vbCrLf & "   </font><font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4</font><p><font size=""2"" face=""Arial""><strong>SECTION 2</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 3</strong></font><br>" & _
                vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 4</strong><br>" & vbCrLf & "   part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font><br>" & _
                vbCrLf & "   <font size=""2"" face=""Arial"">part 6</font><br>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""15""></td>" & vbCrLf & "  <td valign=""top"" width=""80%""><font size=""3""><img src=""" & App.Path & "/Fish.jpg"" width=""188"" height=""103""><br>" & _
                vbCrLf & "   Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat," & _
                " vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font><p><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl " & _
                "ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Contents on Right"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""8"">" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""15""></td>" & vbCrLf & "  <td valign=""bottom"" width=""75%""><font size=""5"" face=""Arial"">Your Heading Goes Here</font></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td width=""25%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & _
                vbCrLf & "  <td width=""15""></td>" & vbCrLf & "  <td valign=""top"" width=""75%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat." & _
                " Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & vbCrLf & "  </td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td valign=""top"" width=""25%""><font size=""2"" face=""Arial""><strong><img src=""" & App.Path & "/Cake.jpg"" width=""131"" height=""100""><br>" & vbCrLf & "   <br>" & vbCrLf & "   SECTION 1</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & _
                vbCrLf & "   part 4</font><p><font size=""2"" face=""Arial""><strong>SECTION 2</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 3</strong></font>" & vbCrLf & "   <br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 4<br>" & vbCrLf & "   </strong>part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 6</font><br>" & vbCrLf & "  </td>" & vbCrLf & "  </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Staggered Sidebar"
            tmp = tmp & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" colspan=""2"" width=""20%""></td>" & vbCrLf & "  <td width=""3%""></td>" & vbCrLf & "  <td valign=""bottom"" width=""64%""><font size=""6"">Your Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""20%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat.Duis autem dolor in hendrerit in vulputate velit esse molestie </font></td>" & vbCrLf & "  <td valign=""top"" width=""20%"" nowrap></td>" & vbCrLf & "  <td valign=""top"" rowspan=""5"" width=""3%""></td>" & _
                vbCrLf & "  <td valign=""top"" rowspan=""5"" width=""64%""><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate " & _
                "velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zril delenit au gue duis dolore te feugat nulla facilisi. </font><p><font size=""3"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font></p>" & _
                vbCrLf & "   <p><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacil isiper suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""20%""></td>" & _
                vbCrLf & "  <td valign=""top"" width=""20%""><font size=""2"" face=""Arial""><img src=""" & App.Path & "/Cake.jpg"" width=""105"" height=""80""><br>" & vbCrLf & "   Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""20%""><font size=""2"" face=""Arial""><img src=""" & App.Path & "/Cake.jpg"" width=""105"" height=""80""><br>" & vbCrLf & "   Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio</font>" & vbCrLf & "  </td>" & vbCrLf & "  <td valign=""top"" width=""20%"" nowrap></td>" & vbCrLf & " </tr>" & _
                vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""20%""></td>" & vbCrLf & "  <td valign=""top"" width=""20%""><font size=""2"" face=""Arial""><img src=""" & App.Path & "/Cake.jpg"" width=""105"" height=""80""><br>" & vbCrLf & "   Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" width=""20%"" nowrap></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Two Sidebars"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""left"">" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""96%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2"" width=""20%""></td>" & vbCrLf & "  <td width=""2%""></td>" & vbCrLf & "  <td colspan=""3""><font size=""7"">Your Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & "  <td valign=""top"" width=""10%""></td>" & _
                vbCrLf & "  <td rowspan=""9"" width=""2%""></td>" & vbCrLf & "  <td valign=""top"" rowspan=""13"" width=""56%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, " & _
                "vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font><p><font size=""2"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font>" & vbCrLf & "  </td>" & vbCrLf & "  <td rowspan=""9"" width=""2%""></td>" & vbCrLf & "  <td align=""right"" valign=""top"" rowspan=""13"" width=""20%""><p align=""right""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""178"" height=""283""><br>" & _
                vbCrLf & "   <font size=""1""><em><strong>Sunflower photo caption</strong></em></font></p>" & vbCrLf & "   <p align=""right""><font size=""2""><em>Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</em></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""></td>" & vbCrLf & "  <td valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & "  <td valign=""top"" width=""10%""></td>" & vbCrLf & " </tr>" & _
                vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""></td>" & vbCrLf & "  <td valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & "  <td valign=""top"" width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""></td>" & vbCrLf & "  <td valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & "  <td valign=""top"" width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & _
                vbCrLf & "  <td align=""right"" valign=""top"" width=""10%""></td>" & _
                vbCrLf & "  <td valign=""top"" width=""10%""><font size=""1"" face=""Arial"">Link to additional information.</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</div>" & vbCrLf & vbCrLf & "<p>&nbsp;</p>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "One-column Body with Two-column Sidebar"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" width=""98%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" width=""60%""><p align=""left""><font size=""7"">Your Heading Goes Here</font></td>" & _
                vbCrLf & "  <td align=""center"" valign=""top""></td>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%"">&nbsp; </td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td valign=""top"" width=""25%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" rowspan=""5"" width=""60%""><p align=""left""><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zril delenit au gue duis dolore te feugat nulla facilisi. </font></p>" & _
                vbCrLf & "   <p align=""left""><font size=""3"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & vbCrLf & "   <p align=""left""><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font></p>" & _
                vbCrLf & "   <p align=""left""><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font></p>" & vbCrLf & "   <p align=""left""><font size=""3"" face=""Arial"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td align=""center"" valign=""top"" rowspan=""5"" width=""4%""></td>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%""><p align=""center""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""86"" height=""112""><strong><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">Topic ONE here</font></strong>" & vbCrLf & "  </td>" & vbCrLf & "  <td rowspan=""5"" width=""4%""></td>" & vbCrLf & "  <td valign=""top"" rowspan=""5"" width=""25%""><p align=""left""><font size=""1"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat.Duis autem dolor in hendrerit in vulputate velit esse molestie </font></p>" & _
                vbCrLf & "   <p align=""left""><font size=""1"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat.</font></p>" & vbCrLf & "   <p align=""left""><font size=""1"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat.</font>" & _
                vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%""><p align=""center""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""86"" height=""112""><br>" & vbCrLf & "   <font size=""2"" face=""Arial""><strong>Topic TWO here</strong></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & _
                vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%""><p align=""center""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""86"" height=""112""><br>" & vbCrLf & "   <font size=""2"" face=""Arial""><strong>Topic THREE here</strong></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""86"" height=""112""><br>" & vbCrLf & "   <font size=""2"" face=""Arial""><strong>Topic Four here</strong></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" valign=""top"" width=""15%""><img src=""" & App.Path & "/Sunflowr.jpg"" width=""86"" height=""112""><br>" & vbCrLf & "   <font size=""2"" face=""Arial""><strong>Topic FIVE here</strong></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Search Page"
            tmp = tmp & vbCrLf & vbCrLf & "<p>Use the form below to search for documents in this web containing specific words or combinations of words. The text search engine will display a weighted list of matching documents, with better matches shown first. Each list item is a link to a matching document; if the document has a title it will be shown, otherwise only the document's file name is displayed. A brief <a href=""#querylang"">explanation</a> of the query language is available, along with examples.</p>" & vbCrLf & vbCrLf & "<form action=""SomeScript.php"" method=""POST"">" & vbCrLf & " Search For : <input type=""text"" name=""SearchString"" size=""35""><br>" & vbCrLf & " <input type=""submit"" name=""Submit"" value=""Start Search"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""reset"" name=""Reset"" value=""Clear"">" & _
                vbCrLf & "</form>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h2><a name=""querylang"">Query Language</a></h2>" & vbCrLf & vbCrLf & "<p>The text search engine allows queries to be formed from arbitrary Boolean expressions containing the keywords AND, OR, and NOT, and grouped with parentheses. For example:</p>" & vbCrLf & vbCrLf & "<blockquote>" & vbCrLf & " <dl>" & _
                vbCrLf & "  <dt><strong><tt>information retrieval</tt></strong></dt>" & vbCrLf & "  <dd>finds documents containing 'information' or 'retrieval'<br><br></dd>" & vbCrLf & "  <dt><strong><tt>information or retrieval</tt></strong></dt>" & vbCrLf & "  <dd>same as above<br><br></dd>" & vbCrLf & "  <dt><strong><tt>information and retrieval</tt></strong></dt>" & vbCrLf & "  <dd>finds documents containing both 'information' and 'retrieval'<br><br></dd>" & vbCrLf & "  <dt><strong><tt>information not retrieval</tt></strong></dt>" & vbCrLf & "  <dd>finds documents containing 'information' but not 'retrieval'<br><br></dd>" & vbCrLf & "  <dt><strong><tt>(information not retrieval) and WAIS</tt></strong></dt>" & vbCrLf & "  <dd>finds documents containing 'WAIS', plus 'information' but not 'retrieval'<br><br></dd>" & vbCrLf & "  <dt><strong><tt>web*</tt></strong></dt>" & vbCrLf & "  <dd>finds documents containing words starting with 'web'<br><br></dd>" & vbCrLf & " </dl>" & vbCrLf & "</blockquote>" & _
                vbCrLf & vbCrLf & "<h5><a href=""#top"">Back to Top</a></h5>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © 1997 Your Company Name. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003 11:17:20.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Table of Contents"
            tmp = tmp & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<p>The following is a hierarchical listing of all the pages in this web that can be reached by following links from the top-level file &quot;index.htm&quot;. Page titles are displayed if they exist, otherwise the entries are file names. Unreachable files are shown at the bottom of the list.</p>" & vbCrLf & vbCrLf & "<p><big><strong><a href=""SomePage.htm"">Table of Contents Heading Page</a></strong></big>" & _
                vbCrLf & vbCrLf & "<ul>" & vbCrLf & " <li><a href=""SomePage.htm"">Title of a Page</a></li>" & vbCrLf & " <li><a href=""SomePage.htm"">Title of a Page</a></li>" & vbCrLf & " <li><a href=""SomePage.htm"">Title of a Page</a></li>" & vbCrLf & "</ul>" & vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © 1997 [OrganizationName]. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003 11:17:20.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Three-column Body"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""90%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" valign=""top"" colspan=""5"" width=""500%""><p align=""center""><font size=""7"">Main Heading Goes Here</font><br>" & _
                vbCrLf & "   <font size=""5"">Subheading Goes Here</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""30%""></td>" & vbCrLf & "  <td width=""5%""></td>" & vbCrLf & "  <td width=""30%""></td>" & vbCrLf & "  <td width=""5%""></td>" & vbCrLf & "  <td width=""30%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & _
                vbCrLf & "  <td valign=""top"" width=""30%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis" & _
                " enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font><p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""5%""></td>" & vbCrLf & "  <td valign=""top"" width=""30%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo " & _
                "consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & vbCrLf & "  </td>" & _
                vbCrLf & "  <td width=""5%""></td>" & vbCrLf & "  <td valign=""top"" width=""30%""><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci" & _
                " taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi" & _
                " per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font><p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo conse</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Two-column Body"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""4"" width=""94%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""bottom"" colspan=""3""><p align=""center""><font size=""6"">Your Heading Goes Here<br>" & vbCrLf & "   </font><font size=""5""><em>Your Section Heading Goes Here</em></font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & _
                vbCrLf & "  <td valign=""top"" width=""50%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod" & _
                " tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis" & _
                " enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font><p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""3%""></td>" & vbCrLf & "  <td valign=""top"" width=""50%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo" & _
                " consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font></p>" & _
                vbCrLf & "   <p><font size=""3"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Two-column Body with Contents on Left"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""8"" width=""98%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""right"" valign=""top"" width=""16%"">&nbsp; </td>" & vbCrLf & "  <td width=""10""></td>" & vbCrLf & "  <td valign=""bottom"" colspan=""3"" width=""60%""><font size=""5"" face=""Arial"">Your Heading Goes Here</font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" width=""16%""><font size=""2"" face=""Arial""><strong>SECTION 1</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4</font><p><font size=""2"" face=""Arial""><strong>SECTION 2</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 3</strong></font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 1<br>" & _
                vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>SECTION 4<br></strong>" & vbCrLf & "   part 1<br>" & vbCrLf & "   part 2<br>" & vbCrLf & "   part 3<br>" & vbCrLf & "   part 4<br>" & vbCrLf & "   part 5</font><br>" & vbCrLf & "   <font size=""2"" face=""Arial"">part 6</font><br>" & vbCrLf & "   <br>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""2%""></td>" & vbCrLf & "  <td valign=""top"" width=""30%""><font size=""3""><img src=""" & App.Path & "/Fish.jpg"" width=""188"" height=""103""><br>" & _
                vbCrLf & "   </font><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""2%""></td>" & _
                vbCrLf & "  <td valign=""top"" width=""30%""><font size=""3"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""3"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat.Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Two-column Body with Two Sidebars"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""6"" width=""96%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""15%""></td>" & vbCrLf & "  <td valign=""top"" colspan=""2""><p align=""center""><font size=""6"" face=""Arial"">Your Heading Goes Here</font></td>" & vbCrLf & "  <td width=""15%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & _
                vbCrLf & "  <td valign=""top"" width=""15%""><font size=""2"" face=""Arial""><strong>Section ONE</strong></font><p><font size=""2"" face=""Arial""><strong>Section TWO</strong></font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section THREE</strong></font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section FOUR</strong></font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section FIVE</strong></font>" & vbCrLf & "  </td>" & _
                vbCrLf & "  <td valign=""top"" width=""30%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font><p><font size=""2"" face=""Arial"">Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat." & _
                "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td valign=""top"" width=""30%""><font size=""2"" face=""Arial"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font><p><font size=""2"" face=""Arial"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""15%""><font size=""1"" face=""Arial"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Two-column Staggered Body"
            tmp = tmp & vbCrLf & vbCrLf & "<p align=""center""><font size=""6""><em>Your Heading Goes Here</em></font></p>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""8"">" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & "  <td align=""right"" width=""40%""><font size=""3""><img src=""" & App.Path & "/earth2.jpg"" width=""88"" height=""85""><br>" & _
                vbCrLf & "   Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & "  <td align=""right"" width=""40%""></td>" & _
                vbCrLf & "  <td width=""40%""><font size=""3""><img src=""" & App.Path & "/Leopard.jpg"" width=""96"" height=""66""><br>" & vbCrLf & "   Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi.</font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & _
                vbCrLf & "  <td align=""right"" width=""40%""><font size=""3""><img src=""" & App.Path & "/Crane.jpg"" width=""95"" height=""114""><br>" & vbCrLf & "   Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. </font>" & _
                vbCrLf & "  </td>" & vbCrLf & "  <td width=""40%""></td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & "  <td align=""right"" width=""40%""></td>" & vbCrLf & "  <td width=""40%""><font size=""3""><img src=""" & App.Path & "/Fish.jpg"" width=""188"" height=""103""><br>" & _
                vbCrLf & "   Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi.</font>" & vbCrLf & "  </td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & _
                vbCrLf & "  <td align=""right"" width=""40%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi.</font></td>" & _
                vbCrLf & "  <td width=""40%""></td>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""10%""></td>" & vbCrLf & "  <td align=""right"" width=""40%""></td>" & _
                vbCrLf & "  <td width=""40%""><font size=""3"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. Ut wisi enim ad minim veniam, quis nostrud exerci taion ullamcorper suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi per suscipit lobortis nisl ut aliquip ex en commodo consequat. Duis te feugifacilisi.</font></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Two-column Staggered Body with Contents and Sidebar"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""96%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" valign=""top"" colspan=""2"" width=""66%""><p align=""center""><font size=""6"" face=""Arial"">Your Heading Goes Here</font></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & _
                vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td valign=""top"" rowspan=""7"" width=""25%""><font size=""2"" face=""Arial""><strong>Section ONE</strong></font><p><font size=""2"" face=""Arial""><strong>Section TWO</strong></font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section THREE</strong></font></p>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section FOUR</strong></font></p>" & vbCrLf & "   <p><font size=""2"" face=""Arial""><strong>Section FIVE</strong></font>" & vbCrLf & "  </td>" & vbCrLf & "  <td></td>" & _
                vbCrLf & "  <td align=""right"" valign=""top"" width=""66%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font></td>" & _
                vbCrLf & "  <td valign=""top"" width=""66%""></td>" & vbCrLf & "  <td rowspan=""7""></td>" & vbCrLf & "  <td valign=""top"" rowspan=""7"" width=""15%""><font size=""1"" face=""Arial"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""></td>" & vbCrLf & "  <td width=""66%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. </font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. </font></td>" & vbCrLf & "  <td width=""66%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""></td>" & vbCrLf & "  <td width=""66%""><font size=""2"" face=""Arial"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim.</font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""><font size=""2"" face=""Arial"">Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & vbCrLf & "  <td width=""66%""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""></td>" & _
                vbCrLf & "  <td width=""66%""><font size=""2"" face=""Arial"">Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td align=""right"" width=""66%""><font size=""2"" face=""Arial"">Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & _
                vbCrLf & "  <td width=""66%""></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "User Registration"
            tmp = tmp & vbCrLf & vbCrLf & "<p>You can automatically register yourself to be a user of [OtherWeb] by filling out and submitting this form. Only registered users are allowed into [OtherWeb]. Choose a username for yourself (such as your last name) and a private password. Together these will be your &quot;key&quot; into [OtherWeb] from now on. This information will be kept in a registration database that is accessible only to the webmaster, not to ordinary users.</p>" & vbCrLf & "<p>One of the main benefits of having a protected web like [OtherWeb] is that authorized users don't have to keep typing their names into form fields, such as when submitting an article to a discussion group, because the web server already knows who they are. Similarly, other users can be reasonably sure that you really sent the articles and postings attributed to you, that someone else didn't pretend to be you when posting.</p>" & _
                vbCrLf & "<p>After you are successfully registered, your web browser will ask you to type in your username and password the first time you try to access [OtherWeb]. The browser will remember this information for as long as it continues to run, so you can access any document in [OtherWeb] without being asked for it again.</p>" & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<form action=""SomeScript.php"" method=""POST"">" & _
                vbCrLf & " <h2>Form Submission</h2>" & vbCrLf & " <p>Make up a username:<br>" & vbCrLf & " <input type=""text"" size=""25"" name=""Username""> -- <em>you can use mixed case</em><br>" & vbCrLf & " Make up a password:<br>" & vbCrLf & " <input type=""password"" size=""25"" name=""Password""> -- <em>keep this private!</em><br>" & vbCrLf & " Enter password again:<br>" & vbCrLf & " <input type=""password"" size=""25"" name=""PasswordVerify""> -- <em>for verification</em><br>" & vbCrLf & " Enter e-mail address:<br>" & vbCrLf & " <input type=""text"" size=""25"" name=""EmailAddress""> -- <em>if you have one</em></p>" & vbCrLf & " <h2><input type=""submit"" value=""Register Me""> <input type=""reset"" value=""Clear Form""></h2>" & vbCrLf & "</form>" & _
                vbCrLf & vbCrLf & "<hr>" & vbCrLf & vbCrLf & "<h5>Author information goes here.<br>" & vbCrLf & "Copyright © 1997 [OrganizationName]. All rights reserved.<br>" & vbCrLf & "Revised: July 04, 2003 11:17:20.</h5>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Wide Body With Headings"
            tmp = tmp & vbCrLf & vbCrLf & "<div align=""center""><center>" & vbCrLf & vbCrLf & "<table border=""0"" width=""80%"">" & vbCrLf & " <tr>" & vbCrLf & "  <td align=""center"" colspan=""2""><font size=""5"" face=""Arial"">Main Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2""></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2""><font size=""4"">Section Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td width=""5%""></td>" & _
                vbCrLf & "  <td width=""60%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2""><font size=""4"">Section Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td width=""60%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font>" & _
                vbCrLf & "   <p><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi.</font>" & vbCrLf & "  </td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2""><font size=""4"">Section Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & _
                vbCrLf & "  <td width=""60%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. Duis autem dolor in hendrerit in vulputate velit esse molestie consequat, illum dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio dignissim qui blandit praesent luptatum zzril delenit au gue duis dolore te feugat nulla facilisi. </font></td>" & _
                vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td colspan=""2""><font size=""4"">Section Heading Goes Here</font></td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td></td>" & vbCrLf & "  <td width=""60%""><font size=""2"" face=""Arial"">Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diem nonummy nibh euismod tincidunt ut lacreet dolore magna aliguam erat volutpat. Ut wisis enim ad minim veniam, quis nostrud exerci tution ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. Duis te feugifacilisi. </font></td>" & vbCrLf & " </tr>" & vbCrLf & "</table>" & vbCrLf & "</center></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
        Case "Banner and Contents"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""64,*"">" & vbCrLf & " <frame name=""banner"" scrolling=""no"" noresize target=""contents"" src=""SomePage1.htm"">" & vbCrLf & " <frameset cols=""150,*"">" & vbCrLf & "  <frame name=""contents"" target=""main"" src=""SomePage2.htm"">" & vbCrLf & "  <frame name=""main"" src=""SomePage3.htm"">" & vbCrLf & " </frameset>" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Contents"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset cols=""150,*"">" & vbCrLf & " <frame name=""contents"" target=""main"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""main"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Footer"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""*,64"">" & vbCrLf & " <frame name=""main"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""footer"" scrolling=""no"" noresize target=""main"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Footnotes"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""*,20%"">" & vbCrLf & " <frame name=""main"" target=""footnotes"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""footnotes"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Header"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""64,*"">" & vbCrLf & " <frame name=""header"" scrolling=""no"" noresize target=""main"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""main"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Header, Footer and Contents"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""64,*,64"">" & vbCrLf & " <frame name=""top"" scrolling=""no"" noresize target=""contents"" src=""SomePage1.htm"">" & vbCrLf & " <frameset cols=""150,*"">" & vbCrLf & "  <frame name=""contents"" target=""main"" src=""SomePage2.htm"">" & vbCrLf & "  <frame name=""main"" src=""SomePage3.htm"">" & vbCrLf & " </frameset>" & vbCrLf & " <frame name=""bottom"" scrolling=""no"" noresize target=""contents"" src=""SomePage4.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Horizontal Split"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""*,*"">" & vbCrLf & " <frame name=""top"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""bottom"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Nested Hierarchy"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset cols=""150,*"">" & vbCrLf & "  <frame name=""left"" scrolling=""no"" noresize target=""rtop"" src=""SomePage1.htm"">" & vbCrLf & " <frameset rows=""20%,*"">" & vbCrLf & "  <frame name=""rtop"" target=""rbottom"" src=""SomePage2.htm"">" & vbCrLf & "  <frame name=""rbottom"" src=""SomePage3.htm"">" & vbCrLf & " </frameset>" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Top-Down Hierarchy"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset rows=""64,25%,*"">" & vbCrLf & " <frame name=""top"" scrolling=""no"" noresize target=""middle"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""middle"" target=""bottom"" src=""SomePage2.htm"">" & vbCrLf & " <frame name=""bottom"" src=""SomePage3.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
        Case "Vertical Split"
            tmp = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentName & " " & PageNumber & "</title>" & vbCrLf & "<meta name=""GENERATOR"" content=""" & ProgName & """>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<frameset cols=""*,*"">" & vbCrLf & " <frame name=""left"" src=""SomePage1.htm"">" & vbCrLf & " <frame name=""right"" src=""SomePage2.htm"">" & vbCrLf & " <noframes>" & vbCrLf & "  <body>" & vbCrLf & "   <p>This page uses frames, but your browser doesn't support them.</p>" & vbCrLf & "  </body>" & vbCrLf & " </noframes>" & vbCrLf & "</frameset>" & vbCrLf & "</html>"
    End Select
    
    Unload Me
    LoadPage tmp, DocumentName & " " & PageNumber
    
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
    
    Set File = fso.GetFile(App.Path & "\Settings")
    
    File.Attributes = Hidden + System
    
    cmdOpen_Click
End Sub

Private Sub TabStrip_Click()
    picExisting.Visible = False
    picNew.Visible = False
    picRecent.Visible = False
    
    If TabStrip.Tabs(1).Selected = True Then picNew.Visible = True
    If TabStrip.Tabs(2).Selected = True Then picExisting.Visible = True
    If TabStrip.Tabs(3).Selected = True Then picRecent.Visible = True
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
