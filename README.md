# Code-VBA-Automatisation-Mail (Tâche : Envoi d'e-mails depuis une feuille "Contacts")

Description
-----------
Un petit macro VBA pour Microsoft Excel qui automatise la création d'e-mails Outlook à partir d'une feuille de contacts. Conçu pour générer des brouillons d'e-mails (ou envoyer directement si modifié) contenant un sujet et un corps multilingue (fr/en), et éventuellement une pièce jointe.

Cas d'usage
----------
Idéal pour envoyer des rapports périodiques (ex. "Crédit Report"), des relances ou tout message répétitif personnalisé par destinataire.

Feuille attendue : "Contacts"
-----------------------------
Le macro lit chaque ligne à partir de la ligne 2. Colonnes attendues :
- Colonne A : langue (ex: `fr` pour français, toute autre valeur -> anglais)
- Colonne B : compte (nom du compte pour le sujet)
- Colonne C : destinataire (adresse e-mail principale)
- Colonne D : cc1 (adresse en copie)
- Colonne E : cc2
- Colonne F : cc3
- Colonne G : commercial (nom)
- Colonne H : mailCommercial (adresse e‑mail du commercial)
- Colonne I : pj (chemin complet vers la pièce jointe, facultatif)

Fonctionnement
-------------
- Le script crée une instance d'Outlook, parcourt chacune des lignes et construit un message.
- Le sujet est : `Crédit Report – <compte> – Avril 2025` (modifiable).
- Le corps est rédigé en français si la colonne A vaut `"fr"`, sinon en anglais.
- Si une pièce jointe est fournie (colonne I), elle est ajoutée.
- Les messages sont ouverts en brouillon (`.Display`). Remplacer par `.Send` pour envoyer automatiquement.

Installation / utilisation
-------------------------
1. Ouvrir le fichier Excel.
2. Menu Développeur → Visual Basic (ou Alt+F11).
3. Insérer un nouveau Module et coller le code VBA.
4. Sauvegarder.
5. Vérifier que la feuille `Contacts` existe et contient les colonnes selon la structure ci‑dessus.
6. Lancer la macro `EnvoiEmails` (via l'éditeur VBA ou raccourci).
7. Autoriser Outlook si une fenêtre d'autorisation apparaît.

Bonnes pratiques et sécurité
----------------------------
- Ne pas utiliser `.Send` avant d'avoir testé sur quelques lignes : `.Send` envoie réellement les emails.
- Ne pas stocker ou partager des informations sensibles dans le fichier sans contrôle d'accès.
- Vérifier les chemins de pièce jointe. Si le fichier n’existe pas, la macro peut échouer.
- Respecter la confidentialité des destinataires (BCC si nécessaire).

Améliorations possibles
-----------------------
- Nettoyage de la liste CC pour n'ajouter que les adresses non vides.
- Verification d'existence du fichier joint avant ajout.
- Ajout d'une colonne "Envoyer" (Oui/Non) pour filtrer les envois.
- Ajout d'un log d'activité (nouvelle feuille "Log").
- Génération dynamique du mois/année dans le sujet.
- Gestion d'erreurs (On Error) et notifications d'échec par ligne.

Exemple rapide d'amélioration (pseudo) :
- Filtrer les CC vides avant de faire `.CC = ...`
- If Len(Trim(cc1)) > 0 Then add cc1 to list
