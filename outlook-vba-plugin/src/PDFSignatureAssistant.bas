' ========================================
' PDF Signature Assistant - VBA Plugin for Outlook
' ========================================
'
' Description: Plugin VBA qui ajoute un bouton "Signer PDFs" dans le ruban Outlook
' Fonction: Signature automatique des PDFs via API localhost:3000
' Auteur: Zouhair Dkhissi
' Date: Juillet 2025
'
' Installation:
' 1. Outlook ? Développeur ? Visual Basic
' 2. Importer ce fichier .bas
' 3. Redémarrer Outlook
' 4. Le bouton "Signer PDFs" apparaît dans le ruban
'
' ========================================

Option Explicit

' Variables globales
Private Const API_BASE_URL As String = "http://localhost:3000"
Private Const TEMP_FOLDER As String = "C:\Temp\PDFSignature\"

' ========================================
' FONCTION PRINCIPALE - Point d'entrée du plugin
' ========================================
Public Sub SignPDFsFromEmail()
    On Error GoTo ErrorHandler

    Dim selectedItem As Object
    Dim mailItem As Outlook.mailItem
    Dim pdfAttachments As Collection
    Dim i As Integer

    Debug.Print "Début macro : vérification explorer actif"
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "? Aucun explorateur Outlook actif.", vbExclamation, "PDF Signature Assistant"
        Exit Sub
    End If

    Debug.Print "Explorer actif OK, vérification sélection"
    Dim selectionCount As Integer
    selectionCount = Application.ActiveExplorer.Selection.count
    Debug.Print "Nb éléments sélectionnés : " & selectionCount

    If selectionCount = 0 Then
        MsgBox "? Veuillez sélectionner un email dans Outlook.", vbExclamation, "PDF Signature Assistant"
        Exit Sub
    End If

    Dim tryMail As Object
    Set tryMail = Nothing
    On Error Resume Next
    Set tryMail = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0

    If tryMail Is Nothing Then
        MsgBox "? Aucun email valide sélectionné.", vbExclamation, "PDF Signature Assistant"
        Debug.Print "tryMail = Nothing après Set"
        Exit Sub
    End If

    Debug.Print "TypeName(tryMail) = " & TypeName(tryMail)
    If TypeName(tryMail) <> "MailItem" Then
        MsgBox "? L'élément sélectionné n'est pas un email.", vbExclamation, "PDF Signature Assistant"
        Debug.Print "TypeName différent de MailItem : " & TypeName(tryMail)
        Exit Sub
    End If

    Debug.Print "Sélection OK, sujet : " & tryMail.Subject

    Set selectedItem = tryMail
    Set mailItem = selectedItem
    Debug.Print "Email sélectionné: " & mailItem.Subject


    
    If TypeName(selectedItem) <> "MailItem" Then
        MsgBox "? Veuillez sélectionner un email dans Outlook.", vbExclamation, "PDF Signature Assistant"
        Exit Sub
    End If
    
    Set mailItem = selectedItem
    
    ' Afficher popup de confirmation
    Dim confirmMsg As String
    confirmMsg = "Email selectionne: " & mailItem.Subject & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "Pieces jointes: " & mailItem.Attachments.count & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "Voulez-vous scanner et signer les PDFs automatiquement ?"
    
    Dim response As VbMsgBoxResult
    response = MsgBox(confirmMsg, vbYesNo + vbQuestion, "PDF Signature Assistant")
    
    If response = vbNo Then Exit Sub
    
    ' Scanner les pièces jointes PDF
    Set pdfAttachments = ScanPDFAttachments(mailItem)
    
    If pdfAttachments.count = 0 Then
        MsgBox "? Aucun fichier PDF trouvé dans cet email.", vbExclamation, "PDF Signature Assistant"
        Exit Sub
    End If
    
    ' Afficher barre de progression
    Dim progressMsg As String
    progressMsg = pdfAttachments.count & " PDF(s) trouve(s)" & vbCrLf & vbCrLf
    progressMsg = progressMsg & "Traitement en cours..." & vbCrLf
    progressMsg = progressMsg & "• Extraction des PDFs" & vbCrLf
    progressMsg = progressMsg & "• Envoi vers API de signature" & vbCrLf
    progressMsg = progressMsg & "• Creation de la reponse" & vbCrLf & vbCrLf
    progressMsg = progressMsg & "Veuillez patienter..."
    
    ' Créer une form de progression (simplifiée avec MsgBox non-bloquant)
    ' Application.StatusBar = "?? Signature des PDFs en cours..."
    
    ' Traiter les PDFs
    Dim signedPDFs As Collection
    Set signedPDFs = ProcessPDFsWithAPI(pdfAttachments, mailItem)
    
    If signedPDFs.count > 0 Then
        ' Créer la réponse automatique
        CreateAutomaticReply mailItem, signedPDFs
        
        ' Message de succès
        Dim successMsg As String
        successMsg = "Traitement termine avec succes !" & vbCrLf & vbCrLf
        successMsg = successMsg & "PDFs signes: " & signedPDFs.count & vbCrLf & vbCrLf
        successMsg = successMsg & "Email de reponse cree automatiquement." & vbCrLf
        successMsg = successMsg & "Ajoutez le destinataire et envoyez !"
        
        MsgBox successMsg, vbInformation, "PDF Signature Assistant"
    Else
        MsgBox "? Erreur lors du traitement des PDFs. Vérifiez l'API sur localhost:3000", vbCritical, "PDF Signature Assistant"
    End If
    
    ' Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    ' Application.StatusBar = False
    MsgBox "? Erreur: " & Err.Description, vbCritical, "PDF Signature Assistant"
End Sub

' ========================================
' SCANNER LES PIECES JOINTES PDF
' ========================================
Private Function ScanPDFAttachments(mailItem As Outlook.mailItem) As Collection
    Dim pdfCollection As New Collection
    Dim attachment As Outlook.attachment
    Dim i As Integer
    
    For i = 1 To mailItem.Attachments.count
        Set attachment = mailItem.Attachments.Item(i)
        
        ' Vérifier l'extension PDF
        If LCase(Right(attachment.fileName, 4)) = ".pdf" Then
            pdfCollection.Add attachment
        End If
    Next i
    
    Set ScanPDFAttachments = pdfCollection
End Function

' ========================================
' TRAITER LES PDFS AVEC L'API
' ========================================
Private Function ProcessPDFsWithAPI(pdfAttachments As Collection, mailItem As Outlook.mailItem) As Collection
    On Error GoTo ErrorHandler

    Dim signedPDFs As Collection
    Dim pdfFiles As New Collection
    Dim i As Integer
    Dim attachment As Outlook.attachment
    Dim tempFilePath As String

    ' Créer le dossier temporaire s'il n'existe pas
    If Dir(TEMP_FOLDER, vbDirectory) = "" Then
        MkDir TEMP_FOLDER
    End If

    ' Sauvegarder les PDFs temporairement et préparer la liste pour l'API
    For i = 1 To pdfAttachments.count
        Set attachment = pdfAttachments.Item(i)
        tempFilePath = TEMP_FOLDER & "temp_" & i & "_" & attachment.fileName
        attachment.SaveAsFile tempFilePath
        pdfFiles.Add Array(attachment.fileName, tempFilePath)
    Next i

    ' >>> C'est ici qu'il faut mettre ces lignes <<<
    Set signedPDFs = CallSignatureAPI(pdfFiles)

    Set ProcessPDFsWithAPI = signedPDFs
    Exit Function

ErrorHandler:
    MsgBox "? Erreur traitement PDFs: " & Err.Description, vbCritical
    Set ProcessPDFsWithAPI = New Collection
End Function

' ========================================
' APPELER L'API DE SIGNATURE (HTTP)
' ========================================
Private Function CallPDFSignatureAPI(filePath As String, fileName As String) As String
    On Error GoTo ErrorHandler
    
    ' NOTE: Cette fonction utilisera XMLHttpRequest pour communiquer avec votre API
    ' Pour l'instant, simulation du traitement
    
    ' Application.StatusBar = "?? Signature IA: " & fileName
    
    ' Simulation d'attente (dans la vraie version, appel HTTP vers localhost:3000)
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' TODO: Implémenter l'appel HTTP réel vers votre API
    ' POST http://localhost:3000/api/process-pdfs-from-outlook
    
    ' Simulation: retourner le chemin du fichier "signé"
    Dim signedPath As String
    signedPath = TEMP_FOLDER & "signed_" & fileName
    
    ' Copier le fichier pour simulation (dans la vraie version: télécharger depuis l'API)
    FileCopy filePath, signedPath
    
    CallPDFSignatureAPI = signedPath
    Exit Function
    
ErrorHandler:
    CallPDFSignatureAPI = ""
End Function

' ========================================
' CREER LA REPONSE AUTOMATIQUE (VERSION CORRIGÉE UTF-8)
' ========================================
Private Sub CreateAutomaticReply(originalMail As Outlook.mailItem, signedPDFs As Collection)
    On Error GoTo ErrorHandler
    
    Dim replyMail As Outlook.mailItem
    Dim replyBody As String
    Dim i As Integer
    Dim pdfInfo As Variant
    
    ' Créer la réponse
    Set replyMail = originalMail.Reply
    
    ' Modifier le sujet (sans caractères spéciaux)
    replyMail.Subject = "RE: " & originalMail.Subject & " - Documents signes"
    
    ' Créer le corps du message (SANS CARACTÈRES ACCENTUÉS)
    replyBody = "Bonjour," & vbCrLf & vbCrLf
    replyBody = replyBody & "Veuillez trouver ci-joint les documents signes :" & vbCrLf & vbCrLf
    
    ' Ajouter la liste des fichiers (avec puces simples)
    For i = 1 To signedPDFs.count
        pdfInfo = signedPDFs.Item(i)
        replyBody = replyBody & "- " & pdfInfo(0) & vbCrLf
    Next i
    
    replyBody = replyBody & vbCrLf & "Les documents ont ete signes electroniquement et sont prets pour utilisation." & vbCrLf & vbCrLf
    replyBody = replyBody & "Cordialement"
    
    ' Définir le corps du message
    replyMail.body = replyBody
    
    ' Attacher les PDFs signés
    For i = 1 To signedPDFs.count
        pdfInfo = signedPDFs.Item(i)
        replyMail.Attachments.Add pdfInfo(1), olByValue, , pdfInfo(0)
        ' Application.StatusBar = "Ajout piece jointe: " & pdfInfo(0)
    Next i
    
    ' Afficher l'email (ne l'envoie pas automatiquement)
    replyMail.Display
    
    ' Nettoyer les fichiers temporaires
    CleanupTempFiles signedPDFs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur creation reponse: " & Err.Description, vbCritical
End Sub

' ========================================
' NETTOYER LES FICHIERS TEMPORAIRES
' ========================================
Private Sub CleanupTempFiles(signedPDFs As Collection)
    On Error Resume Next
    
    Dim i As Integer
    Dim pdfInfo As Variant
    
    ' Supprimer les fichiers temporaires
    For i = 1 To signedPDFs.count
        pdfInfo = signedPDFs.Item(i)
        Kill pdfInfo(1) ' Fichier signé
    Next i
    
    ' Supprimer les fichiers temp_*
    Kill TEMP_FOLDER & "temp_*"
    
    ' Optionnel: supprimer le dossier s'il est vide
    ' RmDir TEMP_FOLDER
End Sub

' ========================================
' FONCTION DE TEST - Vérifier la configuration
' ========================================
Public Sub TestPDFSignatureConfiguration()
    Dim testMsg As String
    
    testMsg = "Test de configuration PDF Signature Assistant" & vbCrLf & vbCrLf
    testMsg = testMsg & "Dossier temporaire: " & TEMP_FOLDER & vbCrLf
    testMsg = testMsg & "API URL: " & API_BASE_URL & vbCrLf & vbCrLf
    testMsg = testMsg & "Plugin VBA charge avec succes !" & vbCrLf
    testMsg = testMsg & "Selectionnez un email avec des PDFs et cliquez sur 'Signer PDFs'"
    
    MsgBox testMsg, vbInformation, "PDF Signature Assistant - Test"
End Sub

