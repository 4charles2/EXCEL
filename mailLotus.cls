VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mailLotus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub mailLotus(ByVal MailDestinataire As String, ByVal Sujet As String, ByVal CorpsMessage As String, _
                ByVal FichierJoint As String)
 
On Error GoTo Gerreur
 
'Dim session As Domino.NotesSession
Dim session As Object
Dim db As Object
Dim doc As Object
Dim attachme As Object
Dim EmbedObj As Object
Dim attachment() As String
Dim i As Integer
 
Set session = CreateObject("notes.notessession")
Set db = session.GETDATABASE("", "")
Call db.openmail
 
Set doc = db.createdocument
 
With doc
    .form = "Memo"
    .sendto = MailDestinataire
    .Subject = Sujet
    .body = CorpsMessage
    .from = session.COMMONUSERNAME
    .posteddate = Now
    .SaveMessageOnSend = True
End With
 
If FichierJoint <> "" Then
    attachment = Split(FichierJoint, ";")
    For i = 0 To UBound(attachment)
        Set attachme = doc.CreateRichTextItem("Attachment")
        Set EmbedObj = attachme.EMBEDOBJECT(1454, "", attachment(i), "Attachment")
    Next i
End If
 'false ou true depend des cas de figure
Call doc.SEND(False)
 
Exit Sub
Gerreur:
MsgBox Err.Number & " : " & Err.Description, vbCritical, "Erreur"
 
End Sub

