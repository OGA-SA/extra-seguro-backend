const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch");
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument } = require("pdf-lib");

const app = express();
const upload = multer();

app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ limit: "20mb", extended: true }));

app.use(express.json());

// ================= ENV =================

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;

const DEFAULT_FOLDER = process.env.FOLDER_PATH || "Extra Seguro";

const allowedOrigins = (process.env.ALLOWED_ORIGIN || "")
  .split(",")
  .filter(Boolean);

// ================= CORS =================

app.use(cors({
  origin: allowedOrigins.length > 0 ? allowedOrigins : "*",
  methods: ["POST","OPTIONS","GET"]
}));

app.options("*", cors());

// ================= SANITY =================

app.get("/", (req,res)=>res.send("✅ Backend funcionando"));


// ================= TOKEN =================

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
  if (!r.ok) throw new Error(JSON.stringify(data));
  return data.access_token;
}


// ================= SHAREPOINT UPLOAD =================

async function uploadToSharePoint(accessToken, buffer, filename, folder) {
  const safeFolder = encodeURI(folder);
  const safeName = encodeURIComponent(filename);

  const uploadUrl =
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${safeFolder}/${safeName}:/content`;

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
    throw new Error(text);
  }

  return res.json();
}


// ================= ENDPOINT ORIGINAL =================

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {

    if (!req.file) {
      return res.status(400).json({ error: "Falta pdf" });
    }

    const filename = req.file.originalname;
    const folder = DEFAULT_FOLDER;

    const token = await getAccessToken();
    const result = await uploadToSharePoint(
      token,
      req.file.buffer,
      filename,
      folder
    );

    res.json({
      ok: true,
      webUrl: result.webUrl,
      name: result.name
    });

  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});


// ================= PDF EDITABLE =================

app.post("/generate-pdf-editable", async (req, res) => {
  try {

    const data = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([600, 800]);
    const form = pdfDoc.getForm();

    // Campos editables ejemplo
    const nombre = form.createTextField("nombre");
    nombre.setText(data.nombre || "");
    nombre.addToPage(page, { x: 50, y: 720, width: 250, height: 24 });

    const detalle = form.createTextField("detalle");
    detalle.setText(data.detalle || "");
    detalle.addToPage(page, { x: 50, y: 670, width: 400, height: 24 });

    const pdfBytes = await pdfDoc.save();
    const buffer = Buffer.from(pdfBytes);

    // ===== subir a SharePoint =====
    const token = await getAccessToken();
    const filename = `editable_${Date.now()}.pdf`;

    const result = await uploadToSharePoint(
      token,
      buffer,
      filename,
      DEFAULT_FOLDER
    );

    res.json({
      ok: true,
      name: filename,
      webUrl: result.webUrl
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});


// ================= START =================

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("🚀 Backend listo puerto", PORT);
});












