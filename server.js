const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch");
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument, StandardFonts } = require("pdf-lib");

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

    const d = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const form = pdfDoc.getForm();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    let y = 800;

    // ===== TITULO =====
    page.drawText("FORMULARIO EXTRA SEGURO", {
      x: 160,
      y,
      size: 16,
      font
    });

    y -= 40;

    // ===== DATOS SUPERIORES =====

    page.drawText("Taller N°:", { x: 40, y, size: 10, font });
    const taller = form.createTextField("taller");
    taller.setText(d.taller || "");
    taller.addToPage(page, { x: 110, y: y-12, width: 120, height: 18 });

    page.drawText("Serie y N°:", { x: 260, y, size: 10, font });
    const serie = form.createTextField("serie");
    serie.setText(d.serieNumero || "");
    serie.addToPage(page, { x: 340, y: y-12, width: 140, height: 18 });

    y -= 30;

    page.drawText("Fecha:", { x: 360, y, size: 10, font });
    const fecha = form.createTextField("fecha");
    fecha.setText(d.fecha || "");
    fecha.addToPage(page, { x: 410, y: y-12, width: 80, height: 18 });

    y -= 30;

    page.drawText("Siniestro:", { x: 360, y, size: 10, font });
    const sin = form.createTextField("siniestro");
    sin.setText(`${d.siniestro1}-${d.siniestro2}`);
    sin.addToPage(page, { x: 430, y: y-12, width: 100, height: 18 });

    y -= 30;

    page.drawText("Dificulta visual:", { x: 360, y, size: 10, font });
    const dv = form.createTextField("visual");
    dv.setText(d.dificultadVisual || "");
    dv.addToPage(page, { x: 460, y: y-12, width: 100, height: 18 });

    // ===== CANVAS =====

    if (d.canvasImage) {
      const base64 = d.canvasImage.split(",")[1];
      const img = await pdfDoc.embedPng(Buffer.from(base64, "base64"));

      page.drawImage(img, {
        x: 40,
        y: y - 200,
        width: 300,
        height: 150
      });
    }

    y -= 220;

    // ===== TABLAS =====

    function drawTabla(tabla, startX) {
      let ty = y;

      tabla.forEach((row, i) => {
        page.drawText(row.pieza || "", { x: startX, y: ty, size: 8, font });

        const c = form.createTextField(`c_${startX}_${i}`);
        c.setText(row.chapa || "");
        c.addToPage(page, { x: startX+120, y: ty-10, width: 40, height: 14 });

        const p = form.createTextField(`p_${startX}_${i}`);
        p.setText(row.pintura || "");
        p.addToPage(page, { x: startX+170, y: ty-10, width: 40, height: 14 });

        ty -= 18;
      });
    }

    drawTabla(d.tabla1 || [], 40);
    drawTabla(d.tabla2 || [], 300);

    // ===== FIRMA =====

    const quien = form.createTextField("quien");
    quien.setText(d.quien || "");
    quien.addToPage(page, { x: 110, y: 68, width: 200, height: 18 });

    // ===== SAVE =====

    const pdfBytes = await pdfDoc.save();
    const buffer = Buffer.from(pdfBytes);

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













