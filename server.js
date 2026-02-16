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

    const data = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const form = pdfDoc.getForm();

    const { StandardFonts, rgb } = require("pdf-lib");
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // ================= HEADER =================

    // Fondo gris título
    page.drawRectangle({
      x: 40,
      y: 800,
      width: 515,
      height: 24,
      color: rgb(0.85,0.85,0.85)
    });

    page.drawText("FORMULARIO EXTRA SEGURO", {
      x: 150,
      y: 806,
      size: 14,
      font: fontBold
    });

    // Imagen cars.jpg (si viene en base64 desde front)
    if(data.logoCars){
      const img = await pdfDoc.embedJpg(Buffer.from(data.logoCars,"base64"));
      page.drawImage(img,{
        x: 40,
        y: 805,
        width: 80,
        height: 20
      });
    }

    // ================= CAMPOS =================

    function field(name, x, y, w=200, h=18, value=""){
      const f = form.createTextField(name);
      f.setText(value || "");
      f.addToPage(page,{ x, y, width:w, height:h });
      f.setFontSize(10);

      page.drawRectangle({
        x, y, width:w, height:h,
        borderWidth:1
      });
    }

    field("taller", 40, 760, 240, 18, data.taller);
    field("serieNumero", 300, 760, 240, 18, data.serieNumero);
    field("fecha", 40, 735, 240, 18, data.fecha);
    field("siniestro", 300, 735, 240, 18, data.siniestro1+"-"+data.siniestro2);

    // ================= PARABRISAS =================

    if(data.canvasImage){
      const img = await pdfDoc.embedPng(
        Buffer.from(data.canvasImage.split(",")[1],"base64")
      );
      page.drawImage(img,{
        x: 40,
        y: 520,
        width: 515,
        height: 180
      });
    }

    // ================= TEXTO ENTRE IMAGEN Y TABLA =================

    page.drawRectangle({ x:40, y:490, width:515, height:36, borderWidth:1 });

    const obs = form.createTextField("dificultadVisual");
    obs.setText(data.dificultadVisual || "");
    obs.addToPage(page,{ x:45, y:495, width:505, height:26 });

    // ================= TABLAS =================

    function drawTabla(tabla, startY){

      const colX = [40, 300, 430];
      const colW = [260, 130, 125];
      let y = startY;

      // encabezados
      ["PIEZA","CHAPA","PINTURA"].forEach((h,i)=>{
        page.drawRectangle({
          x: colX[i],
          y,
          width: colW[i],
          height: 18,
          color: rgb(0.9,0.9,0.9),
          borderWidth:1
        });
        page.drawText(h,{
          x: colX[i]+4,
          y: y+5,
          size:9,
          font: fontBold
        });
      });

      y -= 18;

      tabla.forEach((row,i)=>{
        const vals=[row.pieza,row.chapa,row.pintura];

        vals.forEach((val,c)=>{
          const f = form.createTextField(`t_${startY}_${i}_${c}`);
          f.setText(val || "");
          f.addToPage(page,{
            x: colX[c],
            y,
            width: colW[c],
            height: 18
          });

          page.drawRectangle({
            x: colX[c],
            y,
            width: colW[c],
            height: 18,
            borderWidth:1
          });
        });

        y -= 18;
      });

      return y;
    }

    let y = drawTabla(data.tabla1 || [], 460);
    drawTabla(data.tabla2 || [], y-20);

    // ================= SAVE =================

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














