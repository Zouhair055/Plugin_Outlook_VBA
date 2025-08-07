import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import multer from 'multer';
import { PDFDocument, rgb } from 'pdf-lib';
import fs from 'fs';
import path from 'path';

// IMPORTANT: Charger les variables d'environnement
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// CORS pour Add-in Outlook
app.use(cors({
  origin: ['http://localhost:3001', 'https://localhost:3001', '*'],
  credentials: true
}));

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

// Configuration pour upload de fichiers
const upload = multer({ dest: 'uploads/temp/' });

// Configuration Groq (gratuit)
const GROQ_API_KEY = process.env.GROQ_API_KEY;

// Charger l'image de signature une seule fois au démarrage
let signatureImageBytes = null;
try {
  signatureImageBytes = fs.readFileSync('signn.png');
  console.log('✅ Image de signature chargée');
} catch (error) {
  console.error('❌ Erreur chargement signature:', error.message);
}

// Endpoint de base pour vérifier que l'API fonctionne
app.get('/', (req, res) => {
  res.json({
    message: '🚀 API Signature PDF pour Add-in Outlook',
    status: 'running',
    endpoints: [
      'POST /api/process-pdfs-from-outlook - Traitement complet depuis Outlook',
      'GET /download-signed/:filename - Téléchargement PDFs signés'
    ]
  });
});

// Fonction pour appeler Groq API
async function callGroqAPI(message) {
  if (!GROQ_API_KEY) {
    throw new Error('GROQ_API_KEY non configurée dans le fichier .env');
  }

  console.log('🤖 Appel vers Groq API...');
  
  const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${GROQ_API_KEY}`
    },
    body: JSON.stringify({
      model: 'llama3-70b-8192',
      messages: [{ role: 'user', content: message }],
      temperature: 0.7,
      max_tokens: 2000
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`❌ Erreur Groq API ${response.status}:`, errorText);
    throw new Error(`Erreur Groq API: ${response.status}`);
  }

  const data = await response.json();
  return data.choices[0].message.content;
}

// ENDPOINT PRINCIPAL pour l'Add-in Outlook : Traitement complet des PDFs
app.post('/api/process-pdfs-from-outlook', upload.array('pdfs'), async (req, res) => {
  try {
    console.log('� Traitement PDFs depuis Add-in Outlook...');

    let files = req.files;
    console.log('Body:', {
  pdfs_base64: req.body.pdfs_base64 ? '[base64 data omitted]' : undefined,
  ...Object.fromEntries(Object.entries(req.body).filter(([k]) => k !== 'pdfs_base64'))
});
    console.log('Files:', files);
    // Si aucun fichier reçu, essayer de lire depuis le body (VBA)
    if ((!files || files.length === 0) && req.body && req.body.pdfs_base64) {
      try {
        let pdfsRaw = req.body.pdfs_base64;
        let pdfs;
        try {
          pdfs = JSON.parse(pdfsRaw);
        } catch (e) {
          console.error('Erreur JSON.parse direct:', e, pdfsRaw);
          pdfs = JSON.parse(decodeURIComponent(pdfsRaw));
        }
        files = [];
        for (const pdf of pdfs) {
          const buffer = Buffer.from(pdf.content, 'base64');
          const tempPath = `uploads/temp/${Date.now()}_${pdf.filename}`;
          fs.writeFileSync(tempPath, buffer);
          files.push({
            originalname: pdf.filename,
            path: tempPath
          });
        }
        console.log(`📄 Reçus ${files.length} fichiers PDF en base64 (VBA)`);
      } catch (e) {
        console.error('Erreur décodage pdfs_base64:', e, req.body.pdfs_base64);
        return res.status(400).json({ error: 'Impossible de décoder les PDFs envoyés en base64' });
      }
    }

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'Aucun fichier PDF reçu' });
    }

    console.log(`📄 Traitement de ${files.length} PDF(s)...`);
    const results = [];

    // Traitement de chaque PDF
    for (const file of files) {
      console.log(`📄 Traitement: ${file.originalname}`);

      try {
        // Étape 1: Analyse IA du PDF
        const analysisPrompt = `
Tu es un expert en analyse de documents PDF. Analyse ce document pour déterminer où placer une signature électronique.

Fichier: ${file.originalname}

IMPORTANT: Réponds UNIQUEMENT avec un objet JSON valide, sans texte additionnel.

Format de réponse requis:
{
  "page": 1,
  "x": 420,
  "y": 80,
  "confidence": "high",
  "reasoning": "Zone vide détectée en bas à droite"
}

Règles de placement:
- Page: Généralement page 1 pour les signatures
- X: Entre 350-450 (droite du document A4)
- Y: Entre 50-120 (bas de page)
- Éviter les zones avec du texte
- Privilégier le bas à droite

Réponds SEULEMENT avec le JSON, rien d'autre.
`;

        const analysis = await callGroqAPI(analysisPrompt);
        console.log(`✅ Analyse ${file.originalname}:`, analysis);

        // Étape 2: Application de la signature
        const signedPdf = await applySignatureToPDF(file, analysis);
        
        results.push({
          original: file.originalname,
          signed: signedPdf.filepath,
          downloadUrl: `/download-signed/${path.basename(signedPdf.filepath)}`,
          coordinates: signedPdf.coordinates,
          status: 'success'
        });

        console.log(`✅ ${file.originalname} signé avec succès`);

      } catch (error) {
        console.error(`❌ Erreur traitement ${file.originalname}:`, error.message);
        results.push({
          original: file.originalname,
          error: error.message,
          status: 'error'
        });
      }
    }

    const successfulResults = results.filter(r => r.status === 'success');
    console.log(`🎉 ${successfulResults.length}/${files.length} fichiers traités avec succès`);

    res.json({
      success: true,
      processedFiles: successfulResults,
      totalFiles: files.length,
      successCount: successfulResults.length,
      message: `${successfulResults.length} PDF(s) signé(s) avec succès`
    });

  } catch (error) {
    console.error('❌ Erreur traitement global:', error);
    res.status(500).json({ 
      success: false,
      error: `Erreur lors du traitement: ${error.message}` 
    });
  }
});

// Fonction pour appliquer la signature à un PDF
async function applySignatureToPDF(file, analysisText) {
  if (!signatureImageBytes) {
    throw new Error('Image de signature non disponible');
  }

  // Lire le PDF original
  const pdfBytes = fs.readFileSync(file.path);
  const pdfDoc = await PDFDocument.load(pdfBytes);
  
  // Intégrer l'image de signature
  const signatureImage = await pdfDoc.embedPng(signatureImageBytes);
  
  // Analyser la réponse de Groq pour extraire les coordonnées
  let x = 400, y = 80, page = 0; // Valeurs par défaut
  let confidence = "default";
  let reasoning = "Coordonnées par défaut utilisées";
  
  try {
    // Nettoyer la réponse de Groq
    let cleanAnalysis = analysisText;
    if (cleanAnalysis.includes('```')) {
      cleanAnalysis = cleanAnalysis.replace(/```json\n?/g, '').replace(/```\n?/g, '');
    }
    
    const analysisData = JSON.parse(cleanAnalysis);
    x = analysisData.x || 400;
    y = analysisData.y || 80;
    page = (analysisData.page || 1) - 1; // PDF-lib utilise index 0
    confidence = analysisData.confidence || "default";
    reasoning = analysisData.reasoning || "Coordonnées extraites de l'analyse IA";
    
  } catch (parseError) {
    console.log(`⚠️ Erreur parsing JSON, utilisation des coordonnées par défaut`);
  }

  // Obtenir la page et appliquer la signature
  const pages = pdfDoc.getPages();
  const targetPage = pages[Math.min(page, pages.length - 1)];
  
  targetPage.drawImage(signatureImage, {
    x: x,
    y: y,
    width: 100,
    height: 50,
  });

  // Sauvegarder le PDF signé
  const signedPdfBytes = await pdfDoc.save();
  const timestamp = Date.now();
  const outputPath = `uploads/signed/${timestamp}_${file.originalname}`;
  
  // Créer le dossier s'il n'existe pas
  if (!fs.existsSync('uploads/signed/')) {
    fs.mkdirSync('uploads/signed/', { recursive: true });
  }
  
  fs.writeFileSync(outputPath, signedPdfBytes);

  // Nettoyer le fichier temporaire
  try {
    fs.unlinkSync(file.path);
  } catch (unlinkError) {
    console.log(`⚠️ Erreur suppression fichier temporaire: ${unlinkError.message}`);
  }

  return {
    filepath: outputPath,
    coordinates: { 
      x: x, 
      y: y, 
      page: page + 1,
      confidence: confidence,
      reasoning: reasoning
    }
  };
}

// Endpoint pour télécharger les PDFs signés
app.get('/download-signed/:filename', (req, res) => {
  const filename = req.params.filename;
  const filepath = path.join('uploads/signed', filename);
  
  if (fs.existsSync(filepath)) {
    console.log(`📥 Téléchargement: ${filename}`);
    res.download(filepath);
  } else {
    res.status(404).json({ error: 'Fichier non trouvé' });
  }
});

// Créer les dossiers nécessaires
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads', { recursive: true });
}
if (!fs.existsSync('uploads/temp')) {
  fs.mkdirSync('uploads/temp', { recursive: true });
}
if (!fs.existsSync('uploads/signed')) {
  fs.mkdirSync('uploads/signed', { recursive: true });
}

// Démarrer le serveur
app.listen(PORT, () => {
  console.log(`🚀 API Signature PDF démarrée sur http://localhost:${PORT}`);
  console.log('📧 Spécialement conçue pour Add-in Outlook');
  console.log('🖼️ Image de signature:', signatureImageBytes ? '✅ Chargée' : '❌ Non trouvée');
  
  if (GROQ_API_KEY) {
    console.log('✅ Clé Groq configurée');
  } else {
    console.log('❌ GROQ_API_KEY manquante dans .env');
  }
  
  console.log('🔗 Endpoints disponibles:');
  console.log('   POST /api/process-pdfs-from-outlook - Traitement complet PDFs');
  console.log('   GET  /download-signed/:filename - Téléchargement PDFs signés');
});