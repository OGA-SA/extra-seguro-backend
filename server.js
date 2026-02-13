const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch"); // v2
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");

const app = express();
const upload = multer();

app.use(express.json({ limit: "10mb" }));

// ⚠️ Variables de entorno (configurarlas en Render)
const TENANT_ID = process.env.TENANT_ID;                 
const CLIENT_ID = process.env.CLIENT_ID;                 
const CLIENT_SECRET = process.env.CLIENT_SECRET;         
const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;

// Carpeta por defecto
const DEFAULT_FOLDER = process.env.FOLDER_PATH || "Extra Seguro";

// 🌍 CORS
const allowedOrigins = (process.env.ALLOWED_ORIGIN || "").split(",").filter(Boolean);

app.use(cors({
  origin: allowedOrigins.length > 0 ? allowedOrigins : "*",
  methods: ["POST", "OPTIONS"],
}));

app.options("/upload", cors());
app.options("/generate-pdf-editable", cors());

// Sanity
app.get("/", (req, res) => res.send("✅ Backend funcionando"));


// =========================
// TOKEN GRAPH
// =========================

async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const body = qs.stringify({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const r = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const data = await r.json();
  if (!r.ok) {
    throw new Error(`Token error: ${r.status} - ${JSON.stringify(data)}`);
  }

  return data.access_token;
}


// =========================
// UPLOAD SHAREPOINT
// =========================

async function uploadToSharePoint(accessToken, buffer, filename, folder) {
  const safeFolder = encodeURI(folder);
  const safeName   = encodeURIComponent(filename);

  const uploadUrl  = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${safeFolder}/${safeName}:/content`;

  const res = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/pdf"
    },
    body: buffer
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Error subiendo PDF: ${res.status} - ${text}`);
  }

  return res.json();
}


// =========================
// ENDPOINT ORIGINAL
// =========================

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Falta el archivo 'pdf'" });
    }

    const filename = (req.body.filename || req.file.originalname || "archivo.pdf").trim();
    const folder = (req.body.folder && req.body.folder.trim()) || DEFAULT_FOLDER;

    const accessToken = await getAccessToken();
    const result = await uploadToSharePoint(accessToken, req.file.buffer, filename, folder);

    res.json({
      ok: true,
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      folder: folder,
    });

  } catch (e) {
    console.error("❌ /upload:", e);
    res.status(500).json({ error: e.message });
  }
});


// =========================
// 🆕 PDF EDITABLE
// =========================

app.post("/generate-pdf-editable", async (req, res) => {
  try {

    const data = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const form = pdfDoc.getForm();

    function label(text, x, y){
      page.drawText(text, { x, y, size: 10, font });
    }

    function field(name, x, y, w=200, h=18){
      const f = form.createTextField(name);
      f.addToPage(page, { x, y, width: w, height: h });
      if(data[name]) f.setText(String(data[name]));
    }

    label("Taller:", 50, 800);
    field("taller", 120, 795);

    label("Serie:", 50, 770);
    field("serieNumero", 120, 765);

    label("Siniestro:", 50, 740);
    field("siniestro", 120, 735);

    label("Técnico:", 50, 710);
    field("QUIEN", 120, 705);

    const bytes = await pdfDoc.save();

    const filename = `editable_${Date.now()}.pdf`;

    const token = await getAccessToken();
    const result = await uploadToSharePoint(token, Buffer.from(bytes), filename, DEFAULT_FOLDER);

    res.json({
      ok: true,
      webUrl: result.webUrl
    });

  } catch(e){
    console.error("❌ editable:", e);
    res.status(500).json({ error: e.message });
  }
});


// =========================
// START
// =========================

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`🚀 Backend listo en puerto ${PORT}`);
});









