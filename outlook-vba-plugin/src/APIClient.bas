' ========================================
' API Client - Communication avec l'API PDF Signature
' ========================================
'
' Description: Module pour communiquer avec l'API Node.js sur localhost:3000
' Fonctions: HTTP POST, TÃƒÂ©lÃƒÂ©chargement de fichiers, Gestion d'erreurs
'
' ========================================

Option Explicit

' Constantes pour l'API
' Ajouter aprÃ¨s la ligne 12 (Option Explicit)

' POUR LOCAL (dÃ©veloppement) :
'Private Const API_BASE_URL As String = "http://localhost:3000"

' POUR RENDER (production) :
Private Const API_BASE_URL As String = "https://pdf-signature-api-60nw.onrender.com"

Private Const API_ENDPOINT As String = "/api/process-pdfs-from-outlook"
Private Const DOWNLOAD_ENDPOINT As String = "/download-signed/"
Private Const TIMEOUT_SECONDS As Long = 120

' ========================================
' APPEL HTTP VERS L'API DE SIGNATURE
' ========================================
Public Function CallSignatureAPI(pdfFiles As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Debug.Print "=== DÃ‰BUT CallSignatureAPI ==="
    Debug.Print "Nombre de PDFs Ã  traiter: " & pdfFiles.count
    Debug.Print "Heure dÃ©but API: " & Now

    Dim signedFiles As New Collection
    Dim http As Object
    Dim apiUrl As String
    Dim i As Integer
    Dim pdfFile As Variant
    Dim fileName As String
    Dim filePath As String
    Dim fileBytes() As Byte
    Dim pdfsBase64 As String

    Debug.Print "1. CrÃ©ation objet HTTP..."
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    apiUrl = API_BASE_URL & API_ENDPOINT
    Debug.Print "? URL API: " & apiUrl
    DoEvents

    Debug.Print "2. Construction JSON base64..."
    pdfsBase64 = "["
    For i = 1 To pdfFiles.count
        Debug.Print "Traitement fichier " & i & "/" & pdfFiles.count
        pdfFile = pdfFiles.Item(i)
        fileName = pdfFile(0)
        filePath = pdfFile(1)
        Debug.Print "Fichier: " & fileName & " - Chemin: " & filePath
        
        fileBytes = ReadFileBytes(filePath)
        Debug.Print "Taille fichier: " & UBound(fileBytes) + 1 & " bytes"
        
        If i > 1 Then pdfsBase64 = pdfsBase64 & ","
        pdfsBase64 = pdfsBase64 & "{""filename"":""" & fileName & """,""content"":""" & EncodeBase64(fileBytes) & """}"
        
        DoEvents ' ? CRUCIAL aprÃ¨s chaque fichier
    Next i
    pdfsBase64 = pdfsBase64 & "]"
    Debug.Print "? JSON construit. Taille: " & Len(pdfsBase64) & " caractÃ¨res"
    DoEvents

    Debug.Print "3. POINT CRITIQUE - Avant appel HTTP..."
    Debug.Print "Heure avant HTTP: " & Now
    DoEvents
    
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    
    Debug.Print "4. POINT CRITIQUE - Envoi donnÃ©es..."
    http.Send "{""pdfs_base64"":""" & EscapeJSONString(pdfsBase64) & """}"
    
    Debug.Print "? Envoi terminÃ©. Heure: " & Now
    Debug.Print "HTTP Status: " & http.status
    Debug.Print "Response: " & Left(http.responseText, 200) & "..."
    DoEvents

    If http.status = 200 Then
        Debug.Print "5. Traitement rÃ©ponse..."
        Dim responseText As String
        responseText = http.responseText
        Dim jsonResponse As Object
        Set jsonResponse = ParseJSONResponse(responseText)
        If jsonResponse("success") = True Then
            Set CallSignatureAPI = ProcessSignedFilesResponse(jsonResponse("processedFiles"))
            Debug.Print "? API Success. Fichiers retournÃ©s: " & CallSignatureAPI.count
        Else
            Debug.Print "ERREUR API: " & jsonResponse("error")
            MsgBox "? Erreur API: " & jsonResponse("error"), vbCritical
            Set CallSignatureAPI = New Collection
        End If
    Else
        Debug.Print "ERREUR HTTP: " & http.status & " - " & http.StatusText
        MsgBox "? Erreur HTTP " & http.status & ": " & http.StatusText, vbCritical
        Set CallSignatureAPI = New Collection
    End If
    
    Debug.Print "=== FIN CallSignatureAPI ==="
    Exit Function

ErrorHandler:
    Debug.Print "?? ERREUR CallSignatureAPI: " & Err.Description
    Debug.Print "?? NumÃ©ro: " & Err.Number
    MsgBox "? Erreur communication API: " & Err.Description, vbCritical
    Set CallSignatureAPI = New Collection
End Function

' ========================================
' FONCTION UTILITAIRE POUR Ãƒâ€°CHAPPER LE JSON
' ========================================
Private Function EscapeJSONString(str As String) As String
    Dim s As String
    s = Replace(str, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    EscapeJSONString = s
End Function

' Nouvelle fonction utilitaire pour lire un fichier en bytes
Private Function ReadFileBytes(filePath As String) As Byte()
    Dim fileNum As Integer
    Dim fileLen As Long
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    fileLen = LOF(fileNum)
    If fileLen > 0 Then
        ReDim bytes(0 To fileLen - 1) As Byte
        Get #fileNum, , bytes
    Else
        ReDim bytes(0)
    End If
    Close #fileNum
    ReadFileBytes = bytes
    Exit Function

ErrorHandler:
    MsgBox "? Erreur communication API: " & Err.Description, vbCritical
    ReadFileBytes = ""
    If fileNum > 0 Then Close #fileNum
End Function

' ========================================
' CONSTRUIRE LES DONNEES MULTIPART
' ========================================
Private Function BuildMultipartFormData(pdfFiles As Collection, boundary As String) As String
    Dim formData As String
    Dim pdfFile As Variant
    Dim fileContent As String
    Dim fileName As String
    Dim i As Integer
    
    formData = ""
    
    ' Ajouter chaque fichier PDF
    For i = 1 To pdfFiles.count
        pdfFile = pdfFiles.Item(i)
        fileName = pdfFile(0) ' Nom du fichier
        fileContent = ReadFileAsBase64(CStr(pdfFile(1))) ' Chemin du fichier
        
        formData = formData & "--" & boundary & vbCrLf
        formData = formData & "Content-Disposition: form-data; name=""pdfs""; filename=""" & fileName & """" & vbCrLf
        formData = formData & "Content-Type: application/pdf" & vbCrLf
        formData = formData & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
        formData = formData & fileContent & vbCrLf
        
        Application.StatusBar = "?? PrÃƒÂ©paration: " & fileName & " (" & i & "/" & pdfFiles.count & ")"
    Next i
    
    ' Fermer le multipart
    formData = formData & "--" & boundary & "--" & vbCrLf
    
    BuildMultipartFormData = formData
End Function

' ========================================
' LIRE UN FICHIER EN BASE64
' ========================================
Private Function ReadFileAsBase64(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileContent() As Byte
    Dim fileNum As Integer
    Dim base64String As String
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    
    ReDim fileContent(LOF(fileNum) - 1)
    Get #fileNum, , fileContent
    Close #fileNum
    
    ' Convertir en Base64 (fonction simplifiÃƒÂ©e)
    base64String = EncodeBase64(fileContent)
    
    ReadFileAsBase64 = base64String
    Exit Function
    
ErrorHandler:
    ReadFileAsBase64 = ""
    If fileNum > 0 Then Close #fileNum
End Function

' ========================================
' ENCODER EN BASE64 (VERSION SIMPLIFIEE)
' ========================================
Private Function EncodeBase64(data() As Byte) As String
    ' Version simplifiÃƒÂ©e - dans la vraie implÃƒÂ©mentation,
    ' utiliser MSXML2.DOMDocument ou WinHTTP pour l'encodage Base64
    
    Dim xmlDoc As Object
    Dim dataNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set dataNode = xmlDoc.createElement("data")
    dataNode.DataType = "bin.base64"
    dataNode.nodeTypedValue = data
    
    EncodeBase64 = dataNode.Text
End Function

' ========================================
' PARSER LA REPONSE JSON
' ========================================
Private Function ParseJSONResponse(jsonText As String) As Object
    On Error GoTo ErrorHandler

    Dim json As Object
    Set json = ParseJson(jsonText)
    Set ParseJSONResponse = json
    Exit Function

ErrorHandler:
    Dim errorResult As Object
    Set errorResult = CreateObject("Scripting.Dictionary")
    errorResult("success") = False
    errorResult("error") = "Erreur parsing JSON: " & Err.Description
    Set ParseJSONResponse = errorResult
End Function

' ========================================
' TRAITER LA REPONSE DES FICHIERS SIGNES
' ========================================
Private Function ProcessSignedFilesResponse(processedFiles As Variant) As Collection
    Dim signedFiles As New Collection
    Dim i As Integer
    Dim fileData As Object
    Dim downloadUrl As String
    Dim localPath As String

    On Error GoTo ErrorHandler

    ' GÃƒÂ©rer Array ou Collection
    If IsArray(processedFiles) Then
        For i = LBound(processedFiles) To UBound(processedFiles)
            Debug.Print "Traitement fichier #" & i
            Set fileData = processedFiles(i)
            If Not fileData Is Nothing Then
                On Error Resume Next
                If fileData.Exists("downloadUrl") Then
                    Debug.Print "Appel DownloadSignedFile pour: " & fileData("original")
                    localPath = DownloadSignedFile(downloadUrl, fileData("original"))
                    Debug.Print "RÃƒÂ©sultat DownloadSignedFile: " & localPath
                    If localPath <> "" And Dir(localPath) <> "" And fileLen(localPath) > 0 Then
                        Debug.Print "? Ajout ÃƒÂ  la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non valide pour ajout ÃƒÂ  la collection: " & localPath
                    End If
                End If
                On Error GoTo 0
            End If
        Next i
    ElseIf TypeName(processedFiles) = "Collection" Then
        For i = 1 To processedFiles.count
            Set fileData = processedFiles.Item(i)
            If Not fileData Is Nothing Then
                If fileData.Exists("downloadUrl") Then
                    'downloadUrl = "http://localhost:3000" & fileData("downloadUrl")
                    downloadUrl = API_BASE_URL & fileData("downloadUrl")
                    localPath = DownloadSignedFile(downloadUrl, fileData("original"))
                    If localPath <> "" And Dir(localPath) <> "" And fileLen(localPath) > 0 Then
                        Debug.Print "? Ajout ÃƒÂ  la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non valide pour ajout ÃƒÂ  la collection: " & localPath
                    End If
                ElseIf fileData.Exists("signed") Then
                    Dim signedFileName As String
                    signedFileName = Mid(fileData("signed"), InStrRev(fileData("signed"), "/") + 1)
                    localPath = "C:\Temp\PDFSignature\signed_" & signedFileName
                    If Dir(localPath) <> "" Then
                        Debug.Print "? Fichier trouvÃƒÂ© pour ajout ÃƒÂ  la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non trouvÃƒÂ© pour ajout ÃƒÂ  la collection: " & localPath
                    End If
                End If
            End If
        Next i
    End If
    Debug.Print "Nb fichiers signÃƒÂ©s ajoutÃƒÂ©s: " & signedFiles.count
    Set ProcessSignedFilesResponse = signedFiles
    Exit Function

ErrorHandler:
    Set ProcessSignedFilesResponse = New Collection
End Function

' ========================================
' TELECHARGER UN FICHIER SIGNE (VERSION CORRIGÃƒâ€°E FINALE)
' ========================================
Private Function DownloadSignedFile(downloadUrl As String, fileName As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim localPath As String
    Dim fileNum As Integer
    
    ' CrÃƒÂ©er l'objet HTTP
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Construire l'URL complÃƒÂ¨te si nÃƒÂ©cessaire
    Dim fullUrl As String
    If Left(downloadUrl, 4) = "http" Then
        fullUrl = downloadUrl
    Else
        'fullUrl = "http://localhost:3000" & downloadUrl
        fullUrl = API_BASE_URL & downloadUrl
    End If
    
    Debug.Print "TÃƒÂ©lÃƒÂ©chargement depuis: " & fullUrl
    
    ' Faire la requÃƒÂªte HTTP
    http.Open "GET", fullUrl, False
    http.setRequestHeader "Accept", "application/pdf"
    http.Send
    
    Debug.Print "HTTP Status tÃƒÂ©lÃƒÂ©chargement: " & http.status
    
    If http.status = 200 Then
        ' DÃƒÂ©finir le chemin local
        localPath = "C:\Temp\PDFSignature\signed_" & fileName
        
        ' CrÃƒÂ©er le dossier s'il n'existe pas
        If Dir("C:\Temp\PDFSignature", vbDirectory) = "" Then
            MkDir "C:\Temp\PDFSignature"
        End If
        
        ' Ãƒâ€°crire le fichier en mode binaire
        fileNum = FreeFile
        Open localPath For Binary As #fileNum
        Put #fileNum, 1, http.responseBody
        Close #fileNum
        
        ' VÃƒÂ©rifier que le fichier a ÃƒÂ©tÃƒÂ© crÃƒÂ©ÃƒÂ©
        If Dir(localPath) <> "" And fileLen(localPath) > 0 Then
            Debug.Print "? Fichier tÃƒÂ©lÃƒÂ©chargÃƒÂ©: " & localPath & " - Taille: " & fileLen(localPath)
            DownloadSignedFile = localPath
        Else
            Debug.Print "? Fichier non crÃƒÂ©ÃƒÂ© aprÃƒÂ¨s tÃƒÂ©lÃƒÂ©chargement: " & localPath
            DownloadSignedFile = ""
        End If
    Else
        Debug.Print "? Erreur HTTP tÃƒÂ©lÃƒÂ©chargement: " & http.status & " - " & fullUrl
        DownloadSignedFile = ""
    End If
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Debug.Print "? Erreur DownloadSignedFile: " & Err.Description & " - " & fileName
    DownloadSignedFile = ""
End Function

' ========================================
' TESTER LA CONNEXION API
' ========================================
Public Function TestAPIConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlHttp As Object
    Dim testUrl As String
    
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    ' testUrl = "http://localhost:3000/"
    testUrl = API_BASE_URL & "/"
    
    xmlHttp.Open "GET", testUrl, False
    xmlHttp.Send
    
    If xmlHttp.status = 200 Then
        TestAPIConnection = True
        MsgBox "? Connexion API rÃƒÂ©ussie !" & vbCrLf & "Status: " & xmlHttp.status, vbInformation
    Else
        TestAPIConnection = False
        MsgBox "? Ãƒâ€°chec connexion API" & vbCrLf & "Status: " & xmlHttp.status, vbCritical
    End If
    
    Exit Function
    
ErrorHandler:
    TestAPIConnection = False
    MsgBox "? Erreur test API: " & Err.Description & vbCrLf & vbCrLf & _
           "VÃƒÂ©rifiez que l'API fonctionne sur https://pdf-signature-api-60nw.onrender.com", vbCritical
End Function

Private Function ToBytes(v As Variant) As Byte()
    If VarType(v) = (vbArray + vbByte) Then
        ToBytes = v
    Else
        ReDim ToBytes(0)
    End If
End Function





