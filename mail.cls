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
 
Sub MailActiveSheet()
'Working in Excel 2000-2016
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
 
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
 
    Set Sourcewb = ActiveWorkbook
 
    'Copie la feuille active dans un nouveau classeur
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook
 
    'Determine la version Excel et le type de fichier/format
    With Destwb
        If Val(Application.Version) < 12 Then
            'Quand tu utilise Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'Quand tu utilise 2007-2016
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With
 
    '   Le code ci-dessous change toutes les cellules de la feuille en valeur: Si ça vous interersse
    '    With Destwb.Sheets(1).UsedRange
    '        .Cells.Copy
    '        .Cells.PasteSpecial xlPasteValues
    '        .Cells(1).Select
    '    End With
    '    Application.CutCopyMode = False
 
    'Sauvegare la nouvelle classeur; envoie l'email; efface le classeur créé
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Part of " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
 
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
 
    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .to = "rastabomboclat@developpez.net"
            .CC = "Daan94@developpez.net"
            .BCC = ""
            .Subject = "Teste d'un seul Attachement"
            .Body = "Hi there" & vbCr & "We are testing the sending email via MSExcel spreadsheet"
            .Attachments.Add Destwb.FullName
            'Vous pouvez ajouter d'autres annexes comme ci-dessous
            '.Attachments.Add ("C:\test.txt")
            .Send   'ou utiliser .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With
 
    'Efface le fichier que vous avez envoyé
    Kill TempFilePath & TempFileName & FileExtStr
 
    Set OutMail = Nothing
    Set OutApp = Nothing
 
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
 
    'Une boite de dialogue pour confirmer que l'email a bel et bien été envoyé
    MsgBox Application.UserName & "," & vbCr & "Cette Feuille: " & ActiveSheet.Name & ", a été envoyée par email.", _
    vbOKOnly + vbInformation, ActiveWorkbook.Name & " - Envoie d'email"
 
End Sub

