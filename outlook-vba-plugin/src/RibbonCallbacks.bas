' ========================================
' Ribbon UI - Interface utilisateur personnalisÃƒÂ©e
' ========================================
'
' Description: Configuration du ruban Outlook pour ajouter le bouton "Signer PDFs"
' Fonctions: Callback ribbon, Custom UI XML, Gestion d'ÃƒÂ©vÃƒÂ©nements
'
' ========================================

Option Explicit

' Variable globale pour le ruban
Public MyRibbon As IRibbonUI

' ========================================
' CALLBACK POUR INITIALISER LE RUBAN
' ========================================
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set MyRibbon = ribbon
    Debug.Print "Ã°Å¸â€œÅ’ Ruban PDF Signature Assistant chargÃƒÂ©"
End Sub

' ========================================
' CALLBACK POUR LE BOUTON SIGNATURE
' ========================================
Public Sub OnSignPDFs(control As IRibbonControl)
    ' Appeler la fonction principale de signature
    SignPDFsFromEmail
End Sub

' ========================================
' CALLBACK POUR LE BOUTON TEST API
' ========================================
Public Sub OnTestAPI(control As IRibbonControl)
    ' Tester la connexion ÃƒÂ  l'API
    Dim result As Boolean
    result = TestAPIConnection()
End Sub

' ========================================
' CALLBACK POUR L'ICONE DU BOUTON
' ========================================
Public Sub GetButtonImage(control As IRibbonControl, ByRef image)
    ' Retourner l'icÃƒÂ´ne pour le bouton (optionnel)
    ' Dans ce cas, utilise l'icÃƒÂ´ne par dÃƒÂ©faut d'Outlook
End Sub

' ========================================
' CALLBACK POUR LE LABEL DU BOUTON
' ========================================
Public Sub GetButtonLabel(control As IRibbonControl, ByRef label)
    Select Case control.ID
        Case "btnSignPDFs"
            label = "Signer PDFs"
        Case "btnTestAPI"
            label = "Test API"
        Case Else
            label = "PDF Assistant"
    End Select
End Sub

' ========================================
' CALLBACK POUR LE TOOLTIP
' ========================================
Public Sub GetButtonScreentip(control As IRibbonControl, ByRef screentip)
    Select Case control.ID
        Case "btnSignPDFs"
            screentip = "Signer automatiquement les PDFs de cet email"
        Case "btnTestAPI"
            screentip = "Tester la connexion avec l'API de signature"
        Case Else
            screentip = "Assistant de signature PDF"
    End Select
End Sub

' ========================================
' CALLBACK POUR LA DESCRIPTION LONGUE
' ========================================
Public Sub GetButtonSupertip(control As IRibbonControl, ByRef supertip)
    Select Case control.ID
        Case "btnSignPDFs"
            supertip = "Analyse et signe automatiquement tous les PDFs joints ÃƒÂ  cet email, " & _
                      "puis prÃƒÂ©pare une rÃƒÂ©ponse avec les documents signÃƒÂ©s."
        Case "btnTestAPI"
            supertip = "VÃƒÂ©rifie que l'API de signature fonctionne correctement sur localhost:3000"
        Case Else
            supertip = "Assistant automatique pour la signature de documents PDF"
    End Select
End Sub

' ========================================
' CALLBACK POUR L'ACTIVATION DU BOUTON
' ========================================
Public Sub GetButtonEnabled(control As IRibbonControl, ByRef enabled)
    Dim currentMail As mailItem
    enabled = False
    On Error Resume Next
    If Not Application.ActiveExplorer Is Nothing Then
        If Application.ActiveExplorer.Selection.count > 0 Then
            Set currentMail = Application.ActiveExplorer.Selection.Item(1)
            If Not currentMail Is Nothing Then
                enabled = HasPDFAttachments(currentMail)
            End If
        End If
    End If
    On Error GoTo 0
End Sub

' ========================================
' VERIFIER LA PRESENCE DE PDFS
' ========================================
Private Function HasPDFAttachments(mailItem As mailItem) As Boolean
    Dim att As attachment
    Dim count As Integer
    
    count = 0
    
    For Each att In mailItem.Attachments
        If LCase(Right(att.fileName, 4)) = ".pdf" Then
            count = count + 1
        End If
    Next att
    
    HasPDFAttachments = (count > 0)
End Function

' ========================================
' RAFRAICHIR LE RUBAN
' ========================================
Public Sub RefreshRibbon()
    If Not MyRibbon Is Nothing Then
        MyRibbon.Invalidate
    End If
End Sub

' ========================================
' CALLBACK POUR LES ERREURS DE RUBAN
' ========================================
Public Sub OnRibbonError(ribbon As IRibbonUI, control As IRibbonControl, ByVal fInvalidateControl As Boolean, ByVal strError As String)
    Debug.Print "Ã¢ÂÅ’ Erreur Ruban: " & strError & " (Control: " & control.ID & ")"
    
    ' Log l'erreur pour dÃƒÂ©bogage
    Dim logFile As String
    logFile = Environ("TEMP") & "\PDFSignatureRibbonError.log"
    
    Open logFile For Append As #1
    Print #1, Now & " - Erreur: " & strError & " - Control: " & control.ID
    Close #1
End Sub
