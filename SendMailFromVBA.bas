Attribute VB_Name = "SendMailFromVBA"
Sub SendMail()

Set ObjOutlook = CreateObject("Outlook.Application")


For Linha = 3 To 12

    Set Email = ObjOutlook.CreateItem(0)
    
    Email.Display
    Email.To = Cells(Linha, 1).Value
    Email.CC = Cells(Linha, 2).Value
    Email.BCC = Cells(Linha, 3).Value
    Email.Subject = Cells(Linha, 4).Value
    
    Email.Body = "Olá," & Chr(10) & Chr(10) _
    & Cells(Linha, 5).Value & Chr(10) & Chr(10) _
    & "Atenciosamente," & Chr(10) & "Felipe Bacelo"
    
    Email.Attachments.Add (ThisWorkbook.Path & "\Anexo.xlsx")
    
    Email.Send
    
Next


End Sub
