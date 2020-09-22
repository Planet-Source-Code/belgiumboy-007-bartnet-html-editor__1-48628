Attribute VB_Name = "modStartUp"
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Public Enum ESS
    Yes = 0
    NoSave = 1
    NoDontSave = 2
    NoSaveAll = 3
End Enum

Public Type ESST
    Parameters As ESS
    SaveAllLocation As String
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type typSHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private FileInfo As typSHFILEINFO

Public StartUpScreenShow As Boolean
Public ShowToolBar As Boolean
Public ShowStatusBar As Boolean
Public DocumentName As String
Public DefaultFontType As String
Public DefaultFontSize As Integer
Public DefaultBackGroundColor As String
Public DisableTagColoring As Boolean
Public ExitScreenParameters As ESST
Public HTML_Color As Long
Public Comment_Color As Long

Public Const ProgName As String = "BartNet HTML Editor"
Public Const ProgVersion As String = "1.0.0"

Public PageNumber As Integer
Public FormsCount As Integer
Public frmNewPage As New frmMain
Public Caps(2 To 46) As String

Sub main()
    App.Title = ProgName
    frmMDI.Show
    frmMDI.WindowState = vbMaximized
    frmSplash.Show vbModal, frmMDI
End Sub

Public Sub LoadPage(ByVal Content As String, ByVal Caption As String, Optional NewDocument As Boolean = True, Optional strPath As String = "")
    Set frmNewPage = New frmMain
    
    Load frmNewPage
    
    With frmNewPage
        .Caption = Caption
        .Show
        If frmMDI.ActiveForm.WindowState = vbMaximized Then
            .WindowState = vbMaximized
        Else
            If FormsCount = 0 Then
                .WindowState = vbMaximized
            Else
                .WindowState = vbNormal
            End If
        End If
        
        If NewDocument = True Then
            .Saved = False
            .SavedBefore = False
            .SavedBeforePath = ""
        Else
            .Saved = True
            .SavedBefore = True
            .SavedBeforePath = strPath
        End If
        
        .RichTextBox.Visible = False
        
        .RichTextBox.Text = Content
        
        .RichTextBox.SelStart = 0
        .RichTextBox.SelLength = Len(.RichTextBox.Text)
        .RichTextBox.SelFontSize = DefaultFontSize
        .RichTextBox.SelFontName = DefaultFontType
        .RichTextBox.SelLength = 0
        .RichTextBox.BackColor = DefaultBackGroundColor
        
        .picPreview.Visible = False
        .picEdit.Visible = True
        
        .FirstLoad = True
        .RichTextBox_Change
        .FirstLoad = False
        
        .RichTextBox.Visible = True
    End With
    
    With frmMDI
        Caps(26) = "Preview"
        Set .img(26).Picture = .imgPreview.Picture
        .Toolbar.Buttons("Preview").ToolTipText = "Preview"
        .Toolbar.Buttons("Preview").Image = "Preview"
    End With
    
    FormsCount = FormsCount + 1
End Sub

Public Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Public Function LoadValues() As Boolean
    With frmStartUp
        .picExisting.Visible = False
        .picRecent.Visible = False
    
        .Combo.Text = .Combo.List(0)
    
        .lRecent.ColumnHeaders(2).Width = .lRecent.Width - .lRecent.ColumnHeaders(1).Width - 680
    
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

            Set Item = .lRecent.ListItems.Add(, tmp1, tmp1, ExtractIcon(tmp2 & tmp1, .ilsRecent, .picTmp, 16), ExtractIcon(tmp2 & tmp1, .ilsRecent, .picTmp, 16))
            Item.SubItems(1) = tmp2
        Loop
    
        .lNew.ListItems.Add , "Normal Page", "Normal Page", "Normal", "Normal"
        .lNew.ListItems.Add , "Bibliography", "Bibliography", "Normal", "Normal"
        .lNew.ListItems.Add , "Centered Body", "Centered Body", "Normal", "Normal"
        .lNew.ListItems.Add , "Confirmation Form", "Confirmation Form", "Normal", "Normal"
        .lNew.ListItems.Add , "Feedback Form", "Feedback Form", "Normal", "Normal"
        .lNew.ListItems.Add , "Frequently Asked Questions", "Frequently Asked Questions", "Normal", "Normal"
        .lNew.ListItems.Add , "Guest Book", "Guest Book", "Normal", "Normal"
        .lNew.ListItems.Add , "Narrow, Left-aligned Body", "Narrow, Left-aligned Body", "Normal", "Normal"
        .lNew.ListItems.Add , "Narrow, Right-aligned Body", "Narrow, Right-aligned Body", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Contents and Sidebar", "One-column Body with Contents and Sidebar", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Contents on Left", "One-column Body with Contents on Left", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Contents on Right", "One-column Body with Contents on Right", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Staggered Sidebar", "One-column Body with Staggered Sidebar", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Two Sidebars", "One-column Body with Two Sidebars", "Normal", "Normal"
        .lNew.ListItems.Add , "One-column Body with Two-column Sidebar", "One-column Body with Two-column Sidebar", "Normal", "Normal"
        .lNew.ListItems.Add , "Search Page", "Search Page", "Normal", "Normal"
        .lNew.ListItems.Add , "Table of Contents", "Table of Contents", "Normal", "Normal"
        .lNew.ListItems.Add , "Three-column Body", "Three-column Body", "Normal", "Normal"
        .lNew.ListItems.Add , "Two-column Body", "Two-column Body", "Normal", "Normal"
        .lNew.ListItems.Add , "Two-column Body with Contents on Left", "Two-column Body with Contents on Left", "Normal", "Normal"
        .lNew.ListItems.Add , "Two-column Body with Two Sidebars", "Two-column Body with Two Sidebars", "Normal", "Normal"
        .lNew.ListItems.Add , "Two-column Staggered Body", "Two-column Staggered Body", "Normal", "Normal"
        .lNew.ListItems.Add , "Two-column Staggered Body with Contents and Sidebar", "Two-column Staggered Body with Contents and Sidebar", "Normal", "Normal"
        .lNew.ListItems.Add , "User Registration", "User Registration", "Normal", "Normal"
        .lNew.ListItems.Add , "Wide Body With Headings", "Wide Body With Headings", "Normal", "Normal"
    
        .lNew.ListItems.Add , "Banner and Contents", "Banner and Contents", "Frames", "Frames"
        .lNew.ListItems.Add , "Contents", "Contents", "Frames", "Frames"
        .lNew.ListItems.Add , "Footer", "Footer", "Frames", "Frames"
        .lNew.ListItems.Add , "Footnotes", "Footnotes", "Frames", "Frames"
        .lNew.ListItems.Add , "Header", "Header", "Frames", "Frames"
        .lNew.ListItems.Add , "Header, Footer and Contents", "Header, Footer and Contents", "Frames", "Frames"
        .lNew.ListItems.Add , "Horizontal Split", "Horizontal Split", "Frames", "Frames"
        .lNew.ListItems.Add , "Nested Hierarchy", "Nested Hierarchy", "Frames", "Frames"
        .lNew.ListItems.Add , "Top-Down Hierarchy", "Top-Down Hierarchy", "Frames", "Frames"
        .lNew.ListItems.Add , "Vertical Split", "Vertical Split", "Frames", "Frames"
    
        .lNew.Arrange = lvwAutoLeft
    
        .lNew.ListItems.Item(1).Selected = True
        
        Set Drive = fso.GetDrive("c:\")
        Set Folder2 = fso.GetFolder(Drive.RootFolder)
        
        .tExisting.Nodes.Add , , "Desktop", "Desktop", "Desktop", "Desktop"
        .tExisting.Nodes.Add "Desktop", tvwChild, "My Computer", "My Computer", "MyComputer", "MyComputer"
        
        For Each Drive In fso.Drives
            If Drive.IsReady = True Then
                If Drive.DriveType = Fixed Then
                    .tExisting.Nodes.Add "My Computer", tvwChild, "BEFOREDRIVE" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16), ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16)
                    
                    Set Folder2 = fso.GetFolder(Drive.RootFolder)
                    
                    For Each Folder In Folder2.SubFolders
                        .tExisting.Nodes.Add "BEFOREDRIVE" & Drive.DriveLetter, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, .ilsExisting, .picTmp, 16), ExtractIcon(Folder.Path, .ilsExisting, .picTmp, 16)
                    Next
                    
                    For Each File In Folder2.Files
                        If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                            .tExisting.Nodes.Add "BEFOREDRIVE" & Drive.DriveLetter, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, .ilsExisting, .picTmp, 16), ExtractIcon(File.Path, .ilsExisting, .picTmp, 16)
                        End If
                    Next
                Else
                    .tExisting.Nodes.Add "My Computer", tvwChild, "DRIVE" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16), ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16)
                    
                    Set Folder2 = fso.GetFolder(Drive.RootFolder)
                    
                    For Each Folder In Folder2.SubFolders
                        .tExisting.Nodes.Add "DRIVE" & Drive.DriveLetter, tvwChild, "FOLDER" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, .ilsExisting, .picTmp, 16), ExtractIcon(Folder.Path, .ilsExisting, .picTmp, 16)
                    Next
                    
                    For Each File In Folder2.Files
                        If Right(File.Name, 4) = "html" Or Right(File.Name, 3) = "htm" Then
                            .tExisting.Nodes.Add "DRIVE" & Drive.DriveLetter, tvwChild, "FILE" & File.Path, File.Name, ExtractIcon(File.Path, .ilsExisting, .picTmp, 16), ExtractIcon(File.Path, .ilsExisting, .picTmp, 16)
                        End If
                    Next
                End If
            Else
                .tExisting.Nodes.Add "My Computer", tvwChild, "DRIVE" & Drive.DriveLetter, Drive.DriveLetter, ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16), ExtractIcon(Drive.Path & "\", .ilsExisting, .picTmp, 16)
            End If
        Next
        
        .tExisting.Nodes.Item(1).Expanded = True
        .tExisting.Nodes.Item(2).Expanded = True
        .tExisting.Nodes.Item(4).Expanded = True
    
        Set fso = Nothing
        Set strm = Nothing
        Set Item = Nothing
        Set File = Nothing
        Set Folder = Nothing
        Set Folder2 = Nothing
        Set Drive = Nothing
    End With
    
    LoadValues = True
        
    Exit Function
    
err:
        Set strm = fso.CreateTextFile(App.Path & "\Recent")
        strm.Close
        
        Set File = fso.GetFile(App.Path & "\Recent")
        
        File.Attributes = Hidden + System
        
        Set File = Nothing
        
        LoadValues
End Function

Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
        Case True
            Placement = HWND_TOPMOST
        Case False
            Placement = HWND_NOTOPMOST
    End Select

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
