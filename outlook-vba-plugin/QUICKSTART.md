# 🚀 Quick Start - Plugin VBA PDF Signature Assistant

## Installation Express (5 minutes)

### 1. Démarrer l'API
```powershell
cd "c:\Users\ZouhairDkhissi\OneDrive - All Energies Services\Documents\chat-mistral"
node server.js
```

### 2. Outlook Configuration
- **Ouvrir Outlook**
- **Activer macros:** Fichier → Options → Centre de confidentialité → Paramètres macros → Activer toutes les macros
- **Ouvrir VBA:** `Alt + F11`

### 3. Import Rapide
**Dans l'éditeur VBA:**
```
Fichier → Importer → Sélectionner tous les fichiers .bas du dossier src/
```

### 4. Installation Auto
**Exécuter dans VBA:**
```vba
Sub Install()
    InstallPDFSignaturePlugin
End Sub
```

### 5. Redémarrer Outlook

## ✅ Utilisation Immédiate

1. **Sélectionner un email avec PDFs**
2. **Dans l'éditeur VBA, exécuter:**
   ```vba
   Sub QuickSign()
       SignPDFsFromEmail
   End Sub
   ```
3. **Email prêt avec PDFs signés !**

## 🆘 Dépannage Express

**API ne démarre pas ?**
```powershell
npm install
node server.js
```

**Erreur VBA ?**
- Vérifier macros activées
- Réimporter les modules

**Pas de PDFs ?**
- Vérifier sélection email
- Contrôler pièces jointes PDF

---
*Guide complet: [INSTALLATION.md](INSTALLATION.md)*
