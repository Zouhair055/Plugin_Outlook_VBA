' ========================================
' DANS LE MODULE ThisOutlookSession
' ========================================
Private Sub Application_Startup()
    ' S'exécute automatiquement à chaque démarrage d'Outlook
    AddSignatureButton
    Debug.Print "?? Bouton PDF Signature recréé au démarrage"
End Sub

Private Sub Application_Quit()
    Debug.Print "?? Outlook fermé"
End Sub
