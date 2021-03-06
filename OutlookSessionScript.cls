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
   
    'MsgBox ("D�but de la r�gle PSC")
        
    Set reDate = CreateObject("vbscript.regexp")
    reDate.Pattern = "et la session du [\d]+[\/-][\d]+[\/-][\d]+"
    reDate.Global = True
    Set reMail = CreateObject("vbscript.regexp")
    reMail.Pattern = "\([A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-zA-Z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\)\."
    reMail.Global = True
    Set reName = CreateObject("vbscript.regexp")
    reName.Pattern = "effectu�e par .+ \("
    reName.Global = True

    Debug.Print "D�but de la r�gle PSC : regexs initialis�es"
    
    'Check course date in mail'
    If reDate.Test(Item.Body) Then
        For Each match In reDate.Execute(Item.Body)
            pscDate = Replace(match, "et la session du ", "")
        Next
    Else
        Debug.Print "Pas trouv� de date :("
    End If
    
    'Check mail address in mail'
    If reMail.Test(Item.Body) Then
        For Each matchMail In reMail.Execute(Item.Body)
            pscMail = Replace(Replace(matchMail, "(", ""), ").", "")
        Next
    Else
        Debug.Print "Pas trouv� de mail:("
    End If
    
    'Check name in mail'
    If reName.Test(Item.Body) Then
        For Each matchName In reName.Execute(Item.Body)
            pscName = Replace(Replace(matchName, "effectu�e par ", ""), " (", "")
            
        Next
    Else
        Debug.Print "Pas trouv� de nom:("
    End If
    
    Debug.Print "trouv� ca : " & pscDate
    Debug.Print "trouv� ca : " & pscMail
    Debug.Print "trouv� ca : " & pscName
    

    Set newMail = Application.CreateItem(olMailItem)
    With newMail
        .To = pscMail
        '.CC = "Alias2@domain.com"
        .BCC = "ul.bergesdeloise@croix-rouge.fr"
        .Subject = "Inscription formation PSC1 du " & pscDate
        .BodyFormat = olFormatPlain ' send plain text message
        .Importance = olImportanceHigh
        .Body = "Bonjour " & pscName & " !" & vbCrLf & vbCrLf & "Nous avons bien re�u votre demande d'inscription pour la formation Pr�vention de Secours Civiques de niveau 1 du " & pscDate & " et nous vous en remercions." & vbCrLf & vbCrLf & "Vous trouverez en pi�ce-joint le dossier d'inscription � nous retourner par courrier � l'adresse suivante :" & vbCrLf & vbCrLf & "Croix-Rouge fran�aise" & vbCrLf & "Unit� Locale des Berges de l'Oise - Service formation" & vbCrLf & "41 rue de Villiers-Adam" & vbCrLf & "95290 L'Isle-Adam" & vbCrLf & vbCrLf & "Dans l'attente de vous y rencontrer, nous vous souhaitons une agr�able journ�e et restons � votre enti�re disposition pour toute question." & vbCrLf & vbCrLf & "Cordialement," & vbCrLf & vbCrLf & "L'�quipe formation." & vbCrLf & vbCrLf & vbCrLf & "Croix-Rouge fran�aise" & vbCrLf & "Unit� Locale des Berges de l'Oise" & vbCrLf & "Mail : ul.bergesdeloise@croix-rouge.fr" & vbCrLf & "T�l : 01 34 69 21 99"
        .Attachments.Add ("C:\Dossier_inscription_PSC1.pdf")
    End With
    
    newMail.Send
    Debug.Print "Mail envoy�"
    
    
    Debug.Print "Fin de la r�gle"
End Sub
