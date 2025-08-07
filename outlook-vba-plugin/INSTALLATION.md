# Guide d'Installation - Plugin VBA PDF Signature Assistant

## 📋 Prérequis

- Microsoft Outlook 2010 ou plus récent
- Windows 7 ou plus récent
- API Node.js fonctionnelle sur `http://localhost:3000`
- Macros VBA autorisées dans Outlook

## 🚀 Installation Complète

### Étape 1: Préparation de l'API

1. **Démarrer l'API Node.js:**
   ```powershell
   cd "c:\Users\ZouhairDkhissi\OneDrive - All Energies Services\Documents\chat-mistral"
   node server.js
   ```

2. **Vérifier que l'API fonctionne:**
   - Ouvrir `http://localhost:3000` dans un navigateur
   - Vous devriez voir une page de confirmation

### Étape 2: Configuration d'Outlook

1. **Activer les macros dans Outlook:**
   - Fichier → Options → Centre de gestion de la confidentialité
   - Paramètres du centre de gestion de la confidentialité
   - Paramètres des macros → Activer toutes les macros

2. **Ouvrir l'éditeur VBA:**
   - Dans Outlook: `Alt + F11`
   - Ou Développeur → Visual Basic

### Étape 3: Import des Modules VBA

1. **Importer les modules principaux:**
   ```
   Fichier → Importer un fichier...
   ```
   
   Importer dans cet ordre:
   - `src/PDFSignatureAssistant.bas` (module principal)
   - `src/APIClient.bas` (communication API)
   - `src/RibbonCallbacks.bas` (interface utilisateur)
   - `src/InstallationManager.bas` (installation)

2. **Configuration du ruban personnalisé:**
   - Clic droit sur "Microsoft Outlook Objects"
   - Insérer → Module de classe
   - Copier le contenu de `src/CustomUI.xml`
   - Note: Cette étape nécessite des outils spécialisés pour Outlook

### Étape 4: Installation Automatique

1. **Exécuter l'installation:**
   ```vba
   Sub RunInstallation()
       InstallPDFSignaturePlugin
   End Sub
   ```

2. **Dans l'éditeur VBA:**
   - Sélectionner la fonction `RunInstallation`
   - Appuyer sur `F5` ou cliquer sur ▶️

3. **Suivre les instructions à l'écran**

### Étape 5: Redémarrage et Vérification

1. **Fermer complètement Outlook**
2. **Redémarrer Outlook**
3. **Vérifier l'installation:**
   ```vba
   Sub VerifyInstallation()
       CheckInstallation
   End Sub
   ```

## 🎯 Utilisation Rapide

### Méthode 1: Bouton Ruban (Recommandée)

1. **Sélectionner un email avec PDFs**
2. **Cliquer sur "Signer PDFs"** dans le ruban
3. **Attendre le traitement automatique**
4. **Email de réponse créé automatiquement**

### Méthode 2: Menu Contextuel

1. **Clic droit sur un email**
2. **PDF Signature Assistant → Signer tous les PDFs**

### Méthode 3: Raccourci VBA

1. **Dans l'éditeur VBA:**
   ```vba
   Sub QuickSign()
       SignPDFsFromEmail
   End Sub
   ```

## 🔧 Configuration Avancée

### Fichier de Configuration

Localisation: `C:\Temp\PDFSignature\config.txt`

```ini
[PDF SIGNATURE ASSISTANT - CONFIGURATION]
API_ENDPOINT=http://localhost:3000
SIGNATURE_PATH=
AUTO_REPLY_ENABLED=true
LOG_LEVEL=INFO
BACKUP_ENABLED=true
MAX_FILE_SIZE_MB=50
```

### Personnalisation de la Signature

1. **Modifier l'image de signature:**
   - Remplacer `signn.png` dans le dossier principal
   - Redémarrer l'API Node.js

2. **Ajuster la position:**
   - L'IA Groq détermine automatiquement la position optimale
   - Pour forcer une position, modifier l'API

## 🛠️ Dépannage

### Problème: API non accessible

**Solution:**
1. Vérifier que Node.js est démarré
2. Tester avec: `curl http://localhost:3000`
3. Vérifier le pare-feu Windows

### Problème: Bouton ruban invisible

**Solution:**
1. Vérifier que les macros sont activées
2. Redémarrer Outlook complètement
3. Réimporter les modules VBA

### Problème: Erreur VBA

**Solution:**
1. Ouvrir l'éditeur VBA (`Alt + F11`)
2. Debug → Compiler le projet VBAProject
3. Corriger les erreurs affichées

### Problème: PDFs non signés

**Solution:**
1. Vérifier la clé API Groq dans `server.js`
2. Contrôler les logs: `C:\Temp\PDFSignature\logs\`
3. Tester l'API manuellement

## 📁 Structure des Dossiers

```
C:\Temp\PDFSignature\
├── signed/          # PDFs signés
├── temp/           # Fichiers temporaires
├── logs/           # Journaux d'activité
├── backup/         # Sauvegardes
└── config.txt      # Configuration
```

## 🔄 Mise à Jour

1. **Sauvegarder la configuration actuelle**
2. **Télécharger les nouveaux fichiers VBA**
3. **Réimporter les modules modifiés**
4. **Exécuter à nouveau l'installation**

## 🗑️ Désinstallation

```vba
Sub Uninstall()
    UninstallPDFSignaturePlugin
End Sub
```

1. **Exécuter la fonction de désinstallation**
2. **Supprimer manuellement les modules VBA**
3. **Redémarrer Outlook**

## 🆘 Support

### Logs d'Activité

- **VBA:** `C:\Temp\PDFSignature\logs\`
- **API:** Console Node.js
- **Outlook:** Affichage immédiat dans VBA

### Tests de Diagnostic

```vba
' Test connexion API
Sub TestAPI()
    TestAPIConnection
End Sub

' Vérification complète
Sub FullCheck()
    CheckInstallation
End Sub
```

### Informations de Debug

- **Version VBA:** Vérifier dans `InstallationManager.bas`
- **Version API:** `http://localhost:3000/version`
- **Logs détaillés:** Mode DEBUG dans la configuration

## 📞 Contact

En cas de problème persistant, contacter l'administrateur système avec:

1. **Version d'Outlook**
2. **Message d'erreur exact**
3. **Contenu des fichiers de logs**
4. **Étapes pour reproduire le problème**

---

*Plugin PDF Signature Assistant v1.0 - Développé pour l'intégration Outlook-API*
