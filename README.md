<p align="center">
<a href= "https://img.shields.io/github/repo-size/felipebacelo/SendMailFromVBA?style=for-the-badge"><img src="https://img.shields.io/github/repo-size/felipebacelo/SendMailFromVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/languages/count/felipebacelo/SendMailFromVBA?style=for-the-badge"><img src="https://img.shields.io/github/languages/count/felipebacelo/SendMailFromVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/forks/felipebacelo/SendMailFromVBA?style=for-the-badge"><img src="https://img.shields.io/github/forks/felipebacelo/SendMailFromVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/pr-raw/felipebacelo/SendMailFromVBA?style=for-the-badge"><img src="https://img.shields.io/bitbucket/pr-raw/felipebacelo/SendMailFromVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/issues/felipebacelo/SendMailFromVBA?style=for-the-badge"><img src="https://img.shields.io/bitbucket/issues/felipebacelo/SendMailFromVBA?style=for-the-badge"/></a>
</p>

# SendMailFromVBA

Este respositório foi desenvolvido com o objetivo de automatizar o envio de e-mails utilizando VBA Excel e Outlook.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Outlook 16.0 Object Library

### Compatibilidade

O exemplo deste repositório foi desenvolvido no Excel 2019 (64 bits) e testado no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.
***
### Exemplos de Códigos Utilizados

```vba
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
```
***
### Licenças

_MIT License_
_Copyright   ©   2021 Felipe Bacelo Rodrigues_
