const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch");
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer();

app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ limit: "20mb", extended: true }));

// ================= ENV =================

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;

const DEFAULT_FOLDER = process.env.FOLDER_PATH || "Extra Seguro";
const allowedFolders = ["Formulario 1", "Extra Seguro", "Acta de Restos"];

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

// =====================================================
// ================= ENDPOINT ORIGINAL =================
// =====================================================

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Falta pdf" });
    }

    const filename = req.file.originalname;

    const requestedFolder = req.body.folder?.trim();
    const folder = allowedFolders.includes(requestedFolder)
      ? requestedFolder
      : DEFAULT_FOLDER;

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
      name: result.name,
      folder
    });

  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});


// =====================================================
// ================= PDF EDITABLE ======================
// =====================================================

app.post("/generate-pdf-editable", async (req, res) => {
  try {

    const data = req.body;
    console.log("DATA RECIBIDA:", data);

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const form = pdfDoc.getForm();

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const tableWidth = 240;
    const gap = 20;
    const startX = 40;
    const totalTablesWidth = tableWidth * 2 + gap;

    const logoPath = path.join(__dirname, "cars.jpg");

    if (fs.existsSync(logoPath)) {
      const logoBytes = fs.readFileSync(logoPath);
      const logoImage = await pdfDoc.embedJpg(logoBytes);

      page.drawImage(logoImage,{
        x: 40,
        y: 775,
        width: 90,
        height: 35
      });
    }

    page.drawRectangle({
      x: startX,
      y: 790,
      width: totalTablesWidth,
      height: 20,
      color: rgb(0.9,0.9,0.9)
    });

    const titulo = "FORMULARIO EXTRA SEGURO";
    const tituloWidth = fontBold.widthOfTextAtSize(titulo, 12);

    page.drawText(titulo, {
      x: startX + (totalTablesWidth - tituloWidth) / 2,
      y: 794,
      size: 12,
      font: fontBold,
    });

    function getField(name){
      try {
        return form.getTextField(name);
      } catch {
        return form.createTextField(name);
      }
    }

    function drawLabelField(label, name, x, y, width){
      page.drawRectangle({
        x, y: y - 6,
        width: width + 60,
        height: 20,
        borderWidth: 0.5,
      });

      page.drawText(label,{
        x: x + 4,
        y,
        size: 9,
        font
      });

      const f = getField(name);
      
       f.addToPage(page,{
        x: x + 60,
        y: y - 5,
        width,
        height: 12,
      });
      
      f.setText(data[name] || "");
    }

    drawLabelField("Taller N°:", "taller", startX, 750, 100);
    drawLabelField("Serie y N°:", "serieNumero", startX + 170, 750, 100);
    drawLabelField("Fecha:", "fecha", startX + 340, 750, 80);

    const leftColX = 40;
    const leftColWidth = 330;
    const rightColX = leftColX + leftColWidth + 15;
    const topY = 700;
    const imageHeight = 170;

    if (data.canvasImage && typeof data.canvasImage === "string" && data.canvasImage.includes(",")) {
      const img = await pdfDoc.embedPng(
        Buffer.from(data.canvasImage.split(",")[1],"base64")
      );

      page.drawImage(img,{
        x: leftColX + 1,
        y: topY - imageHeight + 1,
        width: leftColWidth - 2,
        height: imageHeight - 2
      });
    }

    const fila1Y = topY - 20;

    let f;

    f = getField("siniestro");
    f.setText(data.siniestro1 || "");
    f.setFontSize(10);
    f.addToPage(page,{ x: rightColX + 50, y: fila1Y - 5, width: 45, height: 12 });

    f = getField("anio");
    f.setText(data.siniestro2 || "");
    f.setFontSize(10);
    f.addToPage(page,{ x: rightColX + 125, y: fila1Y - 5, width: 35, height: 12 });

    f = getField("dificultadVisual");
    f.setText(data.dificultadVisual || "");
    f.setFontSize(10);
    f.addToPage(page,{
      x: rightColX + 105,
      y: fila1Y - 40,
      width: 120,
      height: 12
    });

    function drawTabla(tabla, startX, startY){
      const colW = [140,50,50];
      const rowH = 16;
      let y = startY;

      (tabla || []).forEach((row,i)=>{
        if (!row) return;

        [row.pieza, row.chapa, row.pintura].forEach((val,c)=>{
          const f = form.createTextField(`tbl_${startX}_${i}_${c}`);
          f.setText(val || "");

          f.addToPage(page,{
            x: startX + colW.slice(0,c).reduce((a,b)=>a+b,0),
            y,
            width: colW[c],
            height: rowH
          });
        });

        y -= rowH;
      });

      return y;
    }

    const tableStartY = 415;
    const bottomTablesY = Math.min(
      drawTabla(data.tabla1, startX, tableStartY),
      drawTabla(data.tabla2, startX + tableWidth + gap, tableStartY)
    );

    const quienField = getField("quien");
    quienField.setText(data.quien || "");
    quienField.addToPage(page,{
      x: startX + 65,
      y: bottomTablesY - 24,
      width: 140,
      height: 14
    });

    
    form.getFields().forEach(field => {
      if (field.setFontSize) {
        field.setFontSize(10);
      }
    });
    
    form.updateFieldAppearances(font);

    const pdfBytes = await pdfDoc.save();

    const token = await getAccessToken();
    const filename = `editable_${Date.now()}.pdf`;

    const result = await uploadToSharePoint(
      token,
      Buffer.from(pdfBytes),
      filename,
      DEFAULT_FOLDER
    );

    res.json({
      ok:true,
      name: filename,
      webUrl: result.webUrl
    });

  } catch (err) {
  
console.error("❌ ERROR PDF EDITABLE:");
  console.error(err);
  console.error(err.stack);

  res.status(500).json({
    error: err.message,
    stack: err.stack });
  }
  });
// ================= START SERVER =================

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log("Servidor corriendo en puerto", PORT);
});















































































