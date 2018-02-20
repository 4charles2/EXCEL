VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 Public bodyMail As String
 Public destinataire As String
 Public CCiDest As String
 Public sujetMail As String
Sub sendMle()
    
    'Fonctionne avec Excel version 2000-2016
    Dim Destwb As Workbook
    Dim OutApp As Object
    Dim OutMail As Object
 
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
 
    'indiquer l'application qui sera utilisé ici Outlook.Application
    Set OutApp = CreateObject("Outlook.Application")
    'Pour note application
    Set OutMail = OutApp.CreateItem(0)
 
    With Destwb
        With OutMail
            .To = destinataire
            .CC = "contact@charles-tognol.fr"
            .BCC = "ctognol@laposte.net"
            .Subject = "Mail depuis excel"
            .body = "Voici mon contenu du mail : " & vbCr & bodyMail
            'Pour ajouter des piece jointe
            '.Attachments.Add Destwb.FullName
            'Vous pouvez ajouter d'autres annexes comme ci-dessous
            '.Attachments.Add ("C:\test.txt")
            .Send   'ou utiliser .Display
        End With
    End With
    'Efface le fichier que vous avez envoyé
 
    Set OutMail = Nothing
    Set OutApp = Nothing
 
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
 
    'Une boite de dialogue pour confirmer que l'email a bel et bien été envoyé
    MsgBox Application.UserName & "," & vbCr & "Les données " & bodyMail & ", a été envoyée par email." & vbCr & " aux destinataire : " & destinataire & " est en copie caché à : " & CCiDest, _
    vbOKOnly + vbInformation, ActiveWorkbook.Name & " - Envoie d'email"
 
End Sub




