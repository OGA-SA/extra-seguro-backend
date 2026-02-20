const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch");
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");

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

    const token = await getAccessToken();
    const result = await uploadToSharePoint(
      token,
      req.file.buffer,
      filename,
      DEFAULT_FOLDER
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


// =====================================================
// ================= PDF EDITABLE ======================
// =====================================================

app.post("/generate-pdf-editable", async (req, res) => {
  try {

    const data = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const form = pdfDoc.getForm();

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // ================= HEADER =================

    // Logo
    if(data.logoCars){
      const img = await pdfDoc.embedJpg(Buffer.from(data.logoCars,"base64"));
      page.drawImage(img,{
        x: 40,
        y: 780,
        width: 90,
        height: 35
      });
    }

    // Fondo gris título
    page.drawRectangle({
      x: 150,
      y: 790,
      width: 300,
      height: 20,
      color: rgb(0.9,0.9,0.9)
    });

    page.drawText("FORMULARIO EXTRA SEGURO", {
      x: 170,
      y: 794,
      size: 12,
      font: fontBold
    });

    // ================= CAMPOS SUPERIORES =================

    function drawLabelField(label, name, x, y, width){
      page.drawText(label,{ x, y: y+4, size:9, font });

      const f = form.createTextField(name);
      f.setText(data[name] || "");
      f.addToPage(page,{ x:x+60, y, width, height:16 });

      page.drawRectangle({
        x:x+60,
        y,
        width,
        height:16,
        borderWidth:0.5
      });
    }

    drawLabelField("Taller N°:", "taller", 40, 750, 100);
    drawLabelField("Serie y N°:", "serieNumero", 210, 750, 100);
    drawLabelField("Fecha:", "fecha", 380, 750, 80);

    drawLabelField("Dificulta visual:", "dificultaVisual", 40, 720, 300);
    drawLabelField("Por cars:", "porCars", 40, 690, 300);

    // ================= TEXTO INFORMATIVO =================

    const infoY = 650;

    page.drawRectangle({ x:40, y:infoY, width:515, height:18, borderWidth:0.5 });
    page.drawText(
      "SE INFORMAN RUBROS CUYOS PORCENTAJES NO SE TENDRAN EN CUENTA EN FUTURAS RECLAMACIONES",
      { x:45, y:infoY+5, size:8, font:fontBold }
    );

    page.drawRectangle({ x:40, y:infoY-18, width:515, height:18, borderWidth:0.5 });
    page.drawText(
      "ANULA/REMPLAZA EXTRA SEGURO DE FECHA:",
      { x:45, y:infoY-13, size:8, font:fontBold }
    );

    const fechaRep = form.createTextField("fechaReemplazo");
    fechaRep.setText(data.fechaReemplazo || "");
    fechaRep.addToPage(page,{
      x:360,
      y:infoY-16,
      width:150,
      height:14
    });

    // ================= TABLAS LADO A LADO =================

    function drawTabla(tabla, startX, startY){

      const colW = [160,70,70];
      const headers = ["PIEZA","CHAPA","PINTURA"];
      const rowH = 16;
      let y = startY;

      headers.forEach((h,i)=>{
        const x = startX + colW.slice(0,i).reduce((a,b)=>a+b,0);

        page.drawRectangle({
          x, y,
          width:colW[i],
          height:rowH,
          color:rgb(0.9,0.9,0.9),
          borderWidth:0.5
        });

        page.drawText(h,{ x:x+4, y:y+4, size:8, font:fontBold });
      });

      y -= rowH;

      (tabla || []).forEach((row,i)=>{
        const vals = [row.pieza, row.chapa, row.pintura];

        vals.forEach((val,c)=>{
          const x = startX + colW.slice(0,c).reduce((a,b)=>a+b,0);

          const f = form.createTextField(`tbl_${startX}_${i}_${c}`);
          f.setText(val || "");
          f.addToPage(page,{ x, y, width:colW[c], height:rowH });

          page.drawRectangle({ x, y, width:colW[c], height:rowH, borderWidth:0.5 });
        });

        y -= rowH;
      });
    }

    drawTabla(data.tabla1, 40, 600);
    drawTabla(data.tabla2, 320, 600);

    // ================= GUARDAR =================

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

  } catch (err){
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});


// ================= START =================

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("🚀 Backend listo puerto", PORT);
});
















