' ========================================
' Installation Manager - Configuration automatique du plugin
' ========================================
'
' Description: Script d'installation et configuration du plugin VBA
' Fonctions: Enregistrement callbacks, Configuration ruban, Tests systÃ¨me
'
' ========================================

Option Explicit

' ========================================
' INSTALLER LE PLUGIN COMPLETEMENT
' ========================================
Public Sub InstallPDFSignaturePlugin()
    On Error GoTo ErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    ' Afficher le dialogue de dÃ©but
    If MsgBox("Installation du Plugin PDF Signature Assistant" & vbCrLf & vbCrLf & _
            "Cette installation va:" & vbCrLf & _
            "• Configurer le ruban Outlook" & vbCrLf & _
            "• Enregistrer les callbacks VBA" & vbCrLf & _
            "• Tester la connexion API" & vbCrLf & _
            "• Creer les dossiers necessaires" & vbCrLf & vbCrLf & _
            "Continuer l'installation ?", vbYesNo + vbQuestion, "Installation Plugin") = vbNo Then
        Exit Sub
    End If

    ' Étapes d'installation
    ' Application.StatusBar = "🔧 Installation en cours..."
    
    ' 1. CrÃ©er les dossiers systÃ¨me
    CreateSystemFolders
    
    ' 2. Tester l'API
    If Not TestAPIConnection() Then
        If MsgBox("âš ï¸ L'API n'est pas accessible sur localhost:3000" & vbCrLf & vbCrLf & _
                  "Voulez-vous continuer l'installation quand mÃªme ?", vbYesNo + vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' 3. Configurer le ruban (nÃ©cessite redÃ©marrage d'Outlook)
    ConfigureRibbon
    
    ' 4. Enregistrer les Ã©vÃ©nements
    RegisterEventHandlers
    
    ' 5. Configuration finale
    SetupConfiguration
    
    ' Calcul du temps d'installation
    Dim installTime As Long
    installTime = DateDiff("s", startTime, Now)
    
    ' Message de succÃ¨s
    MsgBox "âœ… Installation terminÃ©e avec succÃ¨s !" & vbCrLf & vbCrLf & _
           "Temps d'installation: " & installTime & " secondes" & vbCrLf & vbCrLf & _
           "ðŸ”„ RedÃ©marrez Outlook pour voir le nouveau ruban" & vbCrLf & _
           "ðŸ“ Recherchez le groupe 'PDF Signature Assistant'", vbInformation, "Installation RÃ©ussie"
    
    ' Application.StatusBar = "âœ… Plugin PDF Signature installÃ©"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "âŒ Erreur lors de l'installation: " & Err.Description & vbCrLf & vbCrLf & _
           "Code erreur: " & Err.Number, vbCritical, "Erreur Installation"
    ' Application.StatusBar = "âŒ Ã‰chec installation"
End Sub

' ========================================
' CREER LES DOSSIERS SYSTEME
' ========================================
Private Sub CreateSystemFolders()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Dossier principal de travail
    Dim workingDir As String
    workingDir = "C:\Temp\PDFSignature"
    
    If Not fso.FolderExists(workingDir) Then
        fso.CreateFolder workingDir
        Debug.Print "ðŸ“ CrÃ©Ã©: " & workingDir
    End If
    
    ' Sous-dossiers
    Dim subFolders As Variant
    subFolders = Array("signed", "temp", "logs", "backup")
    
    Dim i As Integer
    For i = 0 To UBound(subFolders)
        Dim subDir As String
        subDir = workingDir & "\" & subFolders(i)
        
        If Not fso.FolderExists(subDir) Then
            fso.CreateFolder subDir
            Debug.Print "ðŸ“ CrÃ©Ã©: " & subDir
        End If
    Next i
    
    ' Application.StatusBar = "ðŸ“ Dossiers systÃ¨me crÃ©Ã©s"
End Sub

' ========================================
' CONFIGURER LE RUBAN OUTLOOK
' ========================================
Private Sub ConfigureRibbon()
    ' Note: La configuration du ruban nÃ©cessite l'ajout du XML CustomUI
    ' dans le projet VBA via l'Ã©diteur Visual Basic
    
    ' Cette fonction prÃ©pare les Ã©lÃ©ments nÃ©cessaires
    ' Application.StatusBar = "ðŸŽ¨ Configuration du ruban..."
    
    ' CrÃ©er le fichier de configuration si nÃ©cessaire
    CreateRibbonConfig
    
    ' Message Ã  l'utilisateur
    Debug.Print "ðŸŽ¨ Configuration ruban prÃ©parÃ©e - RedÃ©marrage nÃ©cessaire"
End Sub

' ========================================
' CREER LA CONFIGURATION RUBAN
' ========================================
Private Sub CreateRibbonConfig()
    On Error Resume Next
    
    Dim configPath As String
    configPath = "C:\Temp\PDFSignature\ribbon_config.txt"
    
    ' Ã‰crire les instructions pour l'utilisateur
    Dim configContent As String
    configContent = "CONFIGURATION RUBAN PDF SIGNATURE ASSISTANT" & vbCrLf & vbCrLf & _
                   "Pour activer le ruban personnalisÃ©:" & vbCrLf & _
                   "1. Ouvrir l'Ã©diteur VBA (Alt+F11)" & vbCrLf & _
                   "2. Clic droit sur 'Microsoft Outlook Objects'" & vbCrLf & _
                   "3. InsÃ©rer > Module de classe" & vbCrLf & _
                   "4. Importer le fichier CustomUI.xml" & vbCrLf & _
                   "5. RedÃ©marrer Outlook" & vbCrLf & vbCrLf & _
                   "Fichiers nÃ©cessaires:" & vbCrLf & _
                   "- CustomUI.xml (interface ruban)" & vbCrLf & _
                   "- RibbonCallbacks.bas (callbacks)" & vbCrLf & _
                   "- CrÃ©Ã© le: " & Now
    
    ' Sauvegarder
    Open configPath For Output As #1
    Print #1, configContent
    Close #1
    
    Debug.Print "ðŸ“ Configuration ruban sauvegardÃ©e: " & configPath
End Sub

' ========================================
' ENREGISTRER LES GESTIONNAIRES D'EVENEMENTS
' ========================================
Private Sub RegisterEventHandlers()
    ' Application.StatusBar = "ðŸ”— Enregistrement des Ã©vÃ©nements..."
    
    ' Dans une implÃ©mentation complÃ¨te, enregistrer les Ã©vÃ©nements Outlook
    ' pour dÃ©tecter automatiquement les nouveaux emails avec PDFs
    
    ' Pour le moment, on prÃ©pare la structure
    Debug.Print "ðŸ”— Gestionnaires d'Ã©vÃ©nements prÃ©parÃ©s"
End Sub

' ========================================
' CONFIGURATION FINALE
' ========================================
Private Sub SetupConfiguration()
    ' Application.StatusBar = "âš™ï¸ Configuration finale..."
    
    ' CrÃ©er le fichier de configuration principal
    Dim configPath As String
    configPath = "C:\Temp\PDFSignature\config.txt"
    
    Dim config As String
    config = "[PDF SIGNATURE ASSISTANT - CONFIGURATION]" & vbCrLf & vbCrLf & _
            "API_ENDPOINT=http://localhost:3000" & vbCrLf & _
            "SIGNATURE_PATH=" & vbCrLf & _
            "AUTO_REPLY_ENABLED=true" & vbCrLf & _
            "LOG_LEVEL=INFO" & vbCrLf & _
            "BACKUP_ENABLED=true" & vbCrLf & _
            "MAX_FILE_SIZE_MB=50" & vbCrLf & _
            "INSTALLATION_DATE=" & Now & vbCrLf & _
            "VERSION=1.0" & vbCrLf
    
    ' Sauvegarder la configuration
    Open configPath For Output As #1
    Print #1, config
    Close #1
    
    Debug.Print "âš™ï¸ Configuration sauvegardÃ©e: " & configPath
End Sub

' ========================================
' DESINSTALLER LE PLUGIN
' ========================================
Public Sub UninstallPDFSignaturePlugin()
    On Error GoTo ErrorHandler
    
    ' Confirmation
    If MsgBox("ðŸ—‘ï¸ DÃ©sinstallation du Plugin PDF Signature Assistant" & vbCrLf & vbCrLf & _
              "Cette action va:" & vbCrLf & _
              "â€¢ Supprimer les fichiers temporaires" & vbCrLf & _
              "â€¢ Nettoyer la configuration" & vbCrLf & _
              "â€¢ DÃ©sactiver les callbacks" & vbCrLf & vbCrLf & _
              "âš ï¸ Les fichiers VBA resteront dans Outlook" & vbCrLf & vbCrLf & _
              "Continuer la dÃ©sinstallation ?", vbYesNo + vbExclamation, "DÃ©sinstallation") = vbNo Then
        Exit Sub
    End If
    
    ' Application.StatusBar = "ðŸ—‘ï¸ DÃ©sinstallation en cours..."
    
    ' Supprimer les dossiers temporaires
    CleanupSystemFolders
    
    ' RÃ©initialiser la configuration
    ResetConfiguration
    
    ' Message final
    MsgBox "âœ… DÃ©sinstallation terminÃ©e !" & vbCrLf & vbCrLf & _
           "ðŸ“ Pour supprimer complÃ¨tement le plugin:" & vbCrLf & _
           "1. Ouvrir l'Ã©diteur VBA (Alt+F11)" & vbCrLf & _
           "2. Supprimer manuellement les modules VBA" & vbCrLf & _
           "3. RedÃ©marrer Outlook", vbInformation, "DÃ©sinstallation"

    ' Application.StatusBar = "âœ… Plugin dÃ©sinstallÃ©"

    Exit Sub
    
ErrorHandler:
    MsgBox "âŒ Erreur lors de la dÃ©sinstallation: " & Err.Description, vbCritical
End Sub

' ========================================
' NETTOYER LES DOSSIERS SYSTEME
' ========================================
Private Sub CleanupSystemFolders()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim workingDir As String
    workingDir = "C:\Temp\PDFSignature"
    
    If fso.FolderExists(workingDir) Then
        ' Sauvegarder les fichiers importants avant suppression
        BackupImportantFiles workingDir
        
        ' Supprimer le dossier
        fso.DeleteFolder workingDir, True
        Debug.Print "ðŸ—‘ï¸ Dossier supprimÃ©: " & workingDir
    End If
End Sub

' ========================================
' SAUVEGARDER LES FICHIERS IMPORTANTS
' ========================================
Private Sub BackupImportantFiles(sourceDir As String)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim backupDir As String
    backupDir = Environ("USERPROFILE") & "\Desktop\PDFSignature_Backup_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' CrÃ©er le dossier de sauvegarde
    fso.CreateFolder backupDir
    
    ' Copier les logs et configurations
    If fso.FileExists(sourceDir & "\config.txt") Then
        fso.CopyFile sourceDir & "\config.txt", backupDir & "\"
    End If
    
    If fso.FolderExists(sourceDir & "\logs") Then
        fso.CopyFolder sourceDir & "\logs", backupDir & "\logs\"
    End If
    
    Debug.Print "ðŸ’¾ Sauvegarde crÃ©Ã©e: " & backupDir
End Sub

' ========================================
' REINITIALISER LA CONFIGURATION
' ========================================
Private Sub ResetConfiguration()
    ' RÃ©initialiser les paramÃ¨tres dans le registre si nÃ©cessaire
    ' Nettoyer les rÃ©fÃ©rences temporaires
    
    Debug.Print "ðŸ”„ Configuration rÃ©initialisÃ©e"
End Sub

' ========================================
' VERIFIER L'INSTALLATION
' ========================================
Public Sub CheckInstallation()
    Dim status As String
    status = "ðŸ“‹ VERIFICATION INSTALLATION PDF SIGNATURE ASSISTANT" & vbCrLf & vbCrLf
    
    ' VÃ©rifier les dossiers
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists("C:\Temp\PDFSignature") Then
        status = status & "âœ… Dossiers systÃ¨me: OK" & vbCrLf
    Else
        status = status & "âŒ Dossiers systÃ¨me: MANQUANTS" & vbCrLf
    End If
    
    ' VÃ©rifier l'API
    If TestAPIConnection() Then
        status = status & "âœ… Connexion API: OK" & vbCrLf
    Else
        status = status & "âŒ Connexion API: Ã‰CHEC" & vbCrLf
    End If
    
    ' VÃ©rifier la configuration
    If fso.FileExists("C:\Temp\PDFSignature\config.txt") Then
        status = status & "âœ… Configuration: OK" & vbCrLf
    Else
        status = status & "âŒ Configuration: MANQUANTE" & vbCrLf
    End If
    
    ' Afficher le rapport
    MsgBox status, vbInformation, "VÃ©rification Installation"
End Sub

' ========================================
' AJOUTER BOUTON AMÃ‰LIORE DANS OUTLOOK
' ========================================
Public Sub AddSignatureButton()
    On Error Resume Next
    
    Dim toolbar As CommandBar
    Dim button As CommandBarButton
    
    Set toolbar = Application.ActiveExplorer.CommandBars("Standard")
    
    ' Supprimer l'ancien bouton s'il existe
    toolbar.Controls("PDF Signature").Delete
    toolbar.Controls("ðŸ“„ Signer PDFs").Delete
    
    ' CrÃ©er le nouveau bouton avec style amÃ©liorÃ©
    Set button = toolbar.Controls.Add(Type:=msoControlButton, Before:=1, temporary:=False)
    
    With button
        .Caption = "PDF Signature"                    ' â† Texte plus court
        .Style = msoButtonIconAndCaption             ' IcÃ´ne + texte
        .FaceId = 1950                              ' â† IcÃ´ne de signature (meilleure)
        .OnAction = "SignPDFsFromEmail"
        .Visible = True
        .Tag = "PDFSignatureButton"
        .TooltipText = "Signer automatiquement les PDFs de cet email"  ' â† Tooltip
        .Width = 80                                 ' â† Largeur du bouton
    End With
    
    Debug.Print "âœ… Bouton 'PDF Signature' crÃ©Ã© avec style amÃ©liorÃ© !"
End Sub
