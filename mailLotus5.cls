VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mailLotus5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UseLotus()

    Dim Session As Object
    Dim db As Object
    Dim doc As Object
    Dim rtitem As Object
    Dim object As Object
    Dim fs As Object
    Dim Principaux(2) As String
    Dim Copies(3) As String
    Dim dir As Object
    Dim inti As Integer
    Dim passwd As String
     
    On Error GoTo TraiteErreur
    
    'Demande le password Lotus(Dans le cas ou la session necessite un passwd)
    passwd = InputBox("Entrer votre password Lotus:", "Password")
    
    ' Création de la session Notes
    Set Session = CreateObject("Lotus.NOTESSESSION")
    
    'Ouverture d'une session NOTES
    Call Session.Initialize(passwd) 'si pas de passwd pas de parametre pour initialize

    Set dir = Session.GETDBDIRECTORY("FranceServer1/DCI/BME/Omnia Group")
    Set db = dir.OpenMailDatabase
    
    ' Création d'un document
    Set doc = db.CREATEDOCUMENT

    'affectation du type mail
    Call doc.APPENDITEMVALUE("Form", "Memo")

    Call doc.APPENDITEMVALUE("Sendto", "destinataire@vba.com")
    Call doc.APPENDITEMVALUE("subject", "sujet")
    doc.SAVEMESSAGEONSEND = saveit 'sauvegarde du mail à l envoi
    
    Set rtitem = doc.createRichTextItem("Body")
    
    
    Dim nom As String
    nom = ThisWorkbook.FullName
    'Attachement du classeur au mail
    Set object = rtitem.embedObject(1454, "", nom, "")
    
    Call doc.Send(True)
    Set object = Nothing
    Set rtitem = Nothing
    Set doc = Nothing
    Set db = Nothing
    Set Session = Nothing
    Exit Sub

TraiteErreur:
    MsgBox "Erreur Critique durant l envoi .", vbCritical, "Error"
    Set object = Nothing
    Set rtitem = Nothing
    Set doc = Nothing
    Set db = Nothing
    Set Session = Nothing
    Set fs = Nothing

End Sub
