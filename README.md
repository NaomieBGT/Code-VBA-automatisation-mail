# Code-VBA-automatisation-mail
Sub EnvoiEmails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim xlSheet As Worksheet
    Dim i As Integer
    Dim langue As String, compte As String
    Dim destinataire As String, cc1 As String, cc2 As String, cc3 As String
    Dim commercial As String, mailCommercial As String
    Dim pj As String, objet As String, corps As String

    ' Crée une instance d'Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
   
    ' Sélectionne la feuille "Contacts"
    Set xlSheet = ThisWorkbook.Sheets("Contacts")

    i = 2 ' Commence à la ligne 2 (les données commencent à partir de la ligne 2)

    ' Boucle tant qu'il y a des données dans la colonne A
    Do While xlSheet.Cells(i, 1).Value <> ""
        ' Récupération des données de chaque colonne
        langue = xlSheet.Cells(i, 1).Value
        compte = xlSheet.Cells(i, 2).Value
        destinataire = xlSheet.Cells(i, 3).Value
        cc1 = xlSheet.Cells(i, 4).Value
        cc2 = xlSheet.Cells(i, 5).Value
        cc3 = xlSheet.Cells(i, 6).Value ' Nouvelle colonne cc3 (colonne 6)
        commercial = xlSheet.Cells(i, 7).Value
        mailCommercial = xlSheet.Cells(i, 8).Value
        pj = Replace(xlSheet.Cells(i, 9).Value, """", "") ' Nettoie les guillemets

        ' Création de l'objet (sujet) du mail
        objet = "Crédit Report – " & compte & " – Avril 2025"

        ' Corps du message selon la langue
        If langue = "fr" Then
            corps = "Bonjour," & vbCrLf & vbCrLf & _
                    "Comme prévu, veuillez trouver ci-joint le Crédit Report jusqu'à fin avril 2025." & vbCrLf & vbCrLf & _
                    "Nous avons légèrement réorganisé le document : en haut, vous trouverez un encart regroupant l’ensemble des factures émises ainsi que leurs références. En bas, votre consommation mensuelle y est détaillée mois par mois. Enfin, la colonne située en bas à gauche indique le montant des sessions non encore facturées." & vbCrLf & vbCrLf & _
                    "Bien à vous"
        Else
            corps = "Hello," & vbCrLf & vbCrLf & _
                    "As planned, please find attached the Credit Report up to the end of April 2025." & vbCrLf & vbCrLf & _
                    "We have slightly reorganised the document: at the top, you will find an insert listing all the invoices issued and their references. At the bottom, your monthly consumption is detailed month by month. Finally, the bottom left-hand column shows the amount of sessions not yet billed." & vbCrLf & vbCrLf & _
                    "Kind regards"
        End If

        ' Création et remplissage du mail
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = destinataire
            ' On ajoute cc1, cc2, cc3 et le mail du commercial
            .CC = cc1 & ";" & cc2 & ";" & cc3 & ";" & mailCommercial
            .Subject = objet
            .Body = corps
            If pj <> "" Then .Attachments.Add pj ' Ajoute la pièce jointe si présente
            .Display ' Ouvre le mail sans l’envoyer (utilise .Send pour envoi direct)
        End With

        i = i + 1 ' Passe à la ligne suivante
    Loop

    ' Message de confirmation à la fin
    MsgBox "Tous les mails ont été générés."
End Sub
