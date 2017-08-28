VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub script_psc(Item As Outlook.MailItem)
    Dim re, match
    Dim pj As Outlook.Attachments
    Dim newMail As Outlook.MailItem
   
    'MsgBox ("Début de la règle PSC")
        
    Set reDate = CreateObject("vbscript.regexp")
    reDate.Pattern = "et la session du [\d]+[\/-][\d]+[\/-][\d]+"
    reDate.Global = True
    Set reMail = CreateObject("vbscript.regexp")
    reMail.Pattern = "\([A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-zA-Z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\)\."
    reMail.Global = True
    Set reName = CreateObject("vbscript.regexp")
    reName.Pattern = "effectuée par .+ \("
    reName.Global = True

    Debug.Print "Début de la règle PSC : regexs initialisées"
    
    'Check course date in mail'
    If reDate.Test(Item.Body) Then
        For Each match In reDate.Execute(Item.Body)
            pscDate = Replace(match, "et la session du ", "")
        Next
    Else
        Debug.Print "Pas trouvé de date :("
    End If
    
    'Check mail address in mail'
    If reMail.Test(Item.Body) Then
        For Each matchMail In reMail.Execute(Item.Body)
            pscMail = Replace(Replace(matchMail, "(", ""), ").", "")
        Next
    Else
        Debug.Print "Pas trouvé de mail:("
    End If
    
    'Check name in mail'
    If reName.Test(Item.Body) Then
        For Each matchName In reName.Execute(Item.Body)
            pscName = Replace(Replace(matchName, "effectuée par ", ""), " (", "")
            
        Next
    Else
        Debug.Print "Pas trouvé de nom:("
    End If
    
    Debug.Print "trouvé ca : " & pscDate
    Debug.Print "trouvé ca : " & pscMail
    Debug.Print "trouvé ca : " & pscName
    

    Set newMail = Application.CreateItem(olMailItem)
    With newMail
        .To = pscMail
        '.CC = "Alias2@domain.com"
        .BCC = "ul.bergesdeloise@croix-rouge.fr"
        .Subject = "Inscription formation PSC1 du " & pscDate
        .BodyFormat = olFormatPlain ' send plain text message
        .Importance = olImportanceHigh
        .Body = "Bonjour " & pscName & " !" & vbCrLf & vbCrLf & "Nous avons bien reçu votre demande d'inscription pour la formation Prévention de Secours Civiques de niveau 1 du " & pscDate & " et nous vous en remercions." & vbCrLf & vbCrLf & "Vous trouverez en pièce-joint le dossier d'inscription à nous retourner par courrier à l'adresse suivante :" & vbCrLf & vbCrLf & "Croix-Rouge française" & vbCrLf & "Unité Locale des Berges de l'Oise - Service formation" & vbCrLf & "41 rue de Villiers-Adam" & vbCrLf & "95290 L'Isle-Adam" & vbCrLf & vbCrLf & "Dans l'attente de vous y rencontrer, nous vous souhaitons une agréable journée et restons à votre entière disposition pour toute question." & vbCrLf & vbCrLf & "Cordialement," & vbCrLf & vbCrLf & "L'équipe formation." & vbCrLf & vbCrLf & vbCrLf & "Croix-Rouge française" & vbCrLf & "Unité Locale des Berges de l'Oise" & vbCrLf & "Mail : ul.bergesdeloise@croix-rouge.fr" & vbCrLf & "Tél : 01 34 69 21 99"
        .Attachments.Add ("C:\Dossier_inscription_PSC1.pdf")
    End With
    
    newMail.Send
    Debug.Print "Mail envoyé"
    
    
    Debug.Print "Fin de la règle"
End Sub
