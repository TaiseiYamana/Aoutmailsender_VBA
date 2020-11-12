VBA Code
```
Sub CreateMail()
    If MsgBox("メール作成を実行しますか？", vbOKCancel) = vbCancel Then
        Exit Sub
    End If

    Dim address As Worksheet, mail As Worksheet
    Set address = ThisWorkbook.Sheets("address")
    Set mail = ThisWorkbook.Sheets("メール設定")
    
    Dim OutApp As Outlook.Application 'Outlookアプリケーションオブジェクトを取得
    Set OutApp = New Outlook.Application
    
    Dim r As Long, lastrow As Long
    lastrow = address.Cells(2, 1).End(xlDown).Row
    
    For r = 2 To lastrow
        Dim OutMail As Outlook.MailItem
        Set OutMail = OutApp.CreateItem(olMailItem)
        
        Dim CreatedBody As String '本文作成
        CreatedBody = CreateMailBody(address, mail, r)
        
        Dim CreatedSubject As String '件名作成
        CreatedSubject = CreateMailSubject(address, mail, r)
        
        
        With OutMail
            .SendUsingAccount = Session.Accounts("info@research-p.com")
            .To = address.Cells(r, 5)
            .CC = mail.Cells(1, 2)
            .BCC = mail.Cells(2, 2)
            .subject = CreatedSubject
            .body = CreatedBody
        End With
        OutMail.Display
        Set OutMail = Nothing
        
    Next r
    
End Sub

Function CreateMailBody(address As Worksheet, mail As Worksheet, r As Long) As String
    Dim university As String, lab As String, simei As String, myoji As String, tantou As String '大学名、研究室、氏名､ 苗字、著名
    university = address.Cells(r, 1).Value
    lab = address.Cells(r, 2).Value
    simei = address.Cells(r, 3).Value
    myoji = address.Cells(r, 4).Value
    tantou = mail.Cells(5, 2).Value
    
    Dim body As String
    body = mail.Cells(4, 2).Value
    body = Replace(body, "大学名", university)
    body = Replace(body, "研究室名", lab)
    body = Replace(body, "氏名", simei)
    body = Replace(body, "苗字", myoji)
    body = body & vbCrLf & vbCrLf & tantou '本文＋著名
    
    CreateMailBody = body
End Function

Function CreateMailSubject(address As Worksheet, mail As Worksheet, r As Long) As String
    Dim university As String, lab As String, simei As String, myoji As String '大学名、研究室、氏名、苗字
    university = address.Cells(r, 1).Value
    lab = address.Cells(r, 2).Value
    simei = address.Cells(r, 3).Value
    myoji = address.Cells(r, 4).Value
    
    Dim subject As String
    subject = mail.Cells(3, 2).Value
    subject = Replace(subject, "大学名", university)
    subject = Replace(subject, "研究室名", lab)
    subject = Replace(subject, "氏名", simei)
    subject = Replace(subject, "苗字", myoji)
    
    CreateMailSubject = subject
    
End Function
```
