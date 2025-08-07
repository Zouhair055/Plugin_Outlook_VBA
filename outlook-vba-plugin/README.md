# PDF Signature Assistant - Plugin Outlook VBA

Plugin VBA pour signer automatiquement les PDFs depuis Outlook en un clic.

## 🎯 Fonctionnalités

- **Bouton dans le ruban Outlook** pour signer les PDFs
- **Signature automatique** via IA (Groq)
- **Email de réponse** créé automatiquement avec PDFs signés

## ⚡ Installation

1. **Activer les macros** : Outlook → Options → Centre de confidentialité → Paramètres des macros → Activer
2. **Ouvrir VBA** : `Alt + F11` dans Outlook
3. **Importer les modules** :
   - `src/PDFSignatureAssistant.bas`
   - `src/APIClient.bas` 
   - `src/RibbonCallbacks.bas`
4. **Redémarrer Outlook**

## 🎮 Utilisation

1. Sélectionner un email avec des PDFs
2. Cliquer sur **"PDF Signature"** dans la barre d'outils
3. Confirmer le traitement
4. L'email de réponse s'ouvre avec les PDFs signés

## 🛠️ Dépannage

- **Bouton invisible** : Vérifier que les macros sont activées
- **Erreur API** : Vérifier que l'API fonctionne
- **Logs** : Consulter `C:\Temp\PDFSignature\logs\`

## 🆘 Contact

**Développeur** : Zouhair Dkhissi  
**Email**