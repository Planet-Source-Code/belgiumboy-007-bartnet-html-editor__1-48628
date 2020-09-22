Attribute VB_Name = "modColoring"
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Public Sub ColorTags(rText As RichTextBox, eMidPoint)
    If eMidPoint = 0 Then
        st = 0
        en = Len(rText.Text)
        
        ClearColors rText, 0
    Else
        st = eMidPoint - 100
        
        If st > 0 Then
            Do While Not st = 0
                If Mid(rText.Text, st, 1) = "<" Then Exit Do
                
                st = st - 1
            Loop
        End If
        
        st = st - 1
        
        If st < 0 Then st = 0
        
        en = InStr(eMidPoint, rText.Text, ">")
        If en = 0 Then en = Len(rText.Text)
        
        ClearColors rText, eMidPoint
    End If
    
    While True
        ltTag = rText.Find(Chr(60), st)
        gtTag = rText.Find(Chr(62), ltTag)
        tgLen = (gtTag - ltTag) + 1
        
        If (tgLen) = 0 Or ltTag < 0 Or gtTag < 0 Then Exit Sub
        
        With rText
            .SelStart = ltTag
            .SelLength = tgLen
            
            If Left(.SelText, 2) = "<!" Then .SelColor = Comment_Color Else: .SelColor = HTML_Color
        End With
        
        st = ltTag + tgLen
        If st > en Then Exit Sub
    Wend
End Sub

Public Sub ClearColors(rText As RichTextBox, eMidPoint)
    If eMidPoint = 0 Then
        st = 0
        en = Len(rText.Text)
    Else
        st = eMidPoint - 100
        If st > 0 Then
            Do While Not st = 0
                If Mid(rText.Text, st, 1) = "<" Then _
                Exit Do
                st = st - 1
            Loop
        End If
        st = st - 1
        If st < 0 Then st = 0
        
        en = InStr(eMidPoint, rText.Text, ">")
        If en = 0 Then en = Len(rText.Text)
    End If
    
    With rText
        .SelStart = st
        .SelLength = en - st
        .SelColor = &H0
    End With
End Sub

