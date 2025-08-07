' ========================================
' API Client - Communication avec l'API PDF Signature
' ========================================
'
' Description: Module pour communiquer avec l'API Node.js sur localhost:3000
' Fonctions: HTTP POST, Téléchargement de fichiers, Gestion d'erreurs
'
' ========================================

Option Explicit

' Constantes pour l'API
Private Const API_ENDPOINT As String = "/api/process-pdfs-from-outlook"
Private Const DOWNLOAD_ENDPOINT As String = "/download-signed/"
Private Const TIMEOUT_SECONDS As Long = 120

' ========================================
' APPEL HTTP VERS L'API DE SIGNATURE
' ========================================
Public Function CallSignatureAPI(pdfFiles As Collection) As Collection
    On Error GoTo ErrorHandler

    Dim signedFiles As New Collection
    Dim http As Object
    Dim apiUrl As String
    Dim i As Integer
    Dim pdfFile As Variant
    Dim fileName As String
    Dim filePath As String
    Dim fileBytes() As Byte
    Dim pdfsBase64 As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    apiUrl = "http://localhost:3000/api/process-pdfs-from-outlook"

    ' Construire le JSON base64
    pdfsBase64 = "["
    For i = 1 To pdfFiles.count
        pdfFile = pdfFiles.Item(i)
        fileName = pdfFile(0)
        filePath = pdfFile(1)
        fileBytes = ReadFileBytes(filePath)
        If i > 1 Then pdfsBase64 = pdfsBase64 & ","
        pdfsBase64 = pdfsBase64 & "{""filename"":""" & fileName & """,""content"":""" & EncodeBase64(fileBytes) & """}"
    Next i
    pdfsBase64 = pdfsBase64 & "]"

    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send "{""pdfs_base64"":""" & EscapeJSONString(pdfsBase64) & """}"

    Debug.Print "HTTP Status: " & http.status
    Debug.Print "Response: " & http.responseText

    If http.status = 200 Then
        Dim responseText As String
        responseText = http.responseText
        Dim jsonResponse As Object
        Set jsonResponse = ParseJSONResponse(responseText)
        If jsonResponse("success") = True Then
            Set CallSignatureAPI = ProcessSignedFilesResponse(jsonResponse("processedFiles"))
        Else
            MsgBox "? Erreur API: " & jsonResponse("error"), vbCritical
            Set CallSignatureAPI = New Collection
        End If
    Else
        MsgBox "? Erreur HTTP " & http.status & ": " & http.StatusText, vbCritical
        Set CallSignatureAPI = New Collection
    End If
    Exit Function

ErrorHandler:
    MsgBox "? Erreur communication API: " & Err.Description, vbCritical
    Set CallSignatureAPI = New Collection
End Function

' ========================================
' FONCTION UTILITAIRE POUR ÉCHAPPER LE JSON
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
        
        Application.StatusBar = "?? Préparation: " & fileName & " (" & i & "/" & pdfFiles.count & ")"
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
    
    ' Convertir en Base64 (fonction simplifiée)
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
    ' Version simplifiée - dans la vraie implémentation,
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

    ' Gérer Array ou Collection
    If IsArray(processedFiles) Then
        For i = LBound(processedFiles) To UBound(processedFiles)
            Debug.Print "Traitement fichier #" & i
            Set fileData = processedFiles(i)
            If Not fileData Is Nothing Then
                On Error Resume Next
                If fileData.Exists("downloadUrl") Then
                    Debug.Print "Appel DownloadSignedFile pour: " & fileData("original")
                    localPath = DownloadSignedFile(downloadUrl, fileData("original"))
                    Debug.Print "Résultat DownloadSignedFile: " & localPath
                    If localPath <> "" And Dir(localPath) <> "" And fileLen(localPath) > 0 Then
                        Debug.Print "? Ajout à la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non valide pour ajout à la collection: " & localPath
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
                    downloadUrl = "http://localhost:3000" & fileData("downloadUrl")
                    localPath = DownloadSignedFile(downloadUrl, fileData("original"))
                    If localPath <> "" And Dir(localPath) <> "" And fileLen(localPath) > 0 Then
                        Debug.Print "? Ajout à la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non valide pour ajout à la collection: " & localPath
                    End If
                ElseIf fileData.Exists("signed") Then
                    Dim signedFileName As String
                    signedFileName = Mid(fileData("signed"), InStrRev(fileData("signed"), "/") + 1)
                    localPath = "C:\Temp\PDFSignature\signed_" & signedFileName
                    If Dir(localPath) <> "" Then
                        Debug.Print "? Fichier trouvé pour ajout à la collection: " & localPath
                        signedFiles.Add Array(fileData("original"), localPath, fileData("coordinates"))
                    Else
                        Debug.Print "? Fichier non trouvé pour ajout à la collection: " & localPath
                    End If
                End If
            End If
        Next i
    End If
    Debug.Print "Nb fichiers signés ajoutés: " & signedFiles.count
    Set ProcessSignedFilesResponse = signedFiles
    Exit Function

ErrorHandler:
    Set ProcessSignedFilesResponse = New Collection
End Function

' ========================================
' TELECHARGER UN FICHIER SIGNE (VERSION CORRIGÉE FINALE)
' ========================================
Private Function DownloadSignedFile(downloadUrl As String, fileName As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim localPath As String
    Dim fileNum As Integer
    
    ' Créer l'objet HTTP
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Construire l'URL complète si nécessaire
    Dim fullUrl As String
    If Left(downloadUrl, 4) = "http" Then
        fullUrl = downloadUrl
    Else
        fullUrl = "http://localhost:3000" & downloadUrl
    End If
    
    Debug.Print "Téléchargement depuis: " & fullUrl
    
    ' Faire la requête HTTP
    http.Open "GET", fullUrl, False
    http.setRequestHeader "Accept", "application/pdf"
    http.Send
    
    Debug.Print "HTTP Status téléchargement: " & http.status
    
    If http.status = 200 Then
        ' Définir le chemin local
        localPath = "C:\Temp\PDFSignature\signed_" & fileName
        
        ' Créer le dossier s'il n'existe pas
        If Dir("C:\Temp\PDFSignature", vbDirectory) = "" Then
            MkDir "C:\Temp\PDFSignature"
        End If
        
        ' Écrire le fichier en mode binaire
        fileNum = FreeFile
        Open localPath For Binary As #fileNum
        Put #fileNum, 1, http.responseBody
        Close #fileNum
        
        ' Vérifier que le fichier a été créé
        If Dir(localPath) <> "" And fileLen(localPath) > 0 Then
            Debug.Print "? Fichier téléchargé: " & localPath & " - Taille: " & fileLen(localPath)
            DownloadSignedFile = localPath
        Else
            Debug.Print "? Fichier non créé après téléchargement: " & localPath
            DownloadSignedFile = ""
        End If
    Else
        Debug.Print "? Erreur HTTP téléchargement: " & http.status & " - " & fullUrl
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
    testUrl = "http://localhost:3000/"
    
    xmlHttp.Open "GET", testUrl, False
    xmlHttp.Send
    
    If xmlHttp.status = 200 Then
        TestAPIConnection = True
        MsgBox "? Connexion API réussie !" & vbCrLf & "Status: " & xmlHttp.status, vbInformation
    Else
        TestAPIConnection = False
        MsgBox "? Échec connexion API" & vbCrLf & "Status: " & xmlHttp.status, vbCritical
    End If
    
    Exit Function
    
ErrorHandler:
    TestAPIConnection = False
    MsgBox "? Erreur test API: " & Err.Description & vbCrLf & vbCrLf & _
           "Vérifiez que l'API fonctionne sur http://localhost:3000", vbCritical
End Function

Private Function ToBytes(v As Variant) As Byte()
    If VarType(v) = (vbArray + vbByte) Then
        ToBytes = v
    Else
        ReDim ToBytes(0)
    End If
End Function


