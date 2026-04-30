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

app.use(cors({
  origin: allowedOrigins.length > 0 ? allowedOrigins : "*",
  methods: ["POST", "OPTIONS", "GET"]
}));

app.options("*", cors());

app.get("/", (req, res) => res.send("Backend funcionando"));

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
  let stage = "inicio";

  try {
    const data = req.body || {};
    console.log("DATA RECIBIDA PDF EDITABLE:", Object.keys(data));

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);
    const form = pdfDoc.getForm();

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const black = rgb(0, 0, 0);
    const gray = rgb(0.8, 0.8, 0.8);
    const borderGray = rgb(0.6, 0.6, 0.6);

    function clean(value) {
      if (value === undefined || value === null) return "";
      return String(value)
        .replace(/[^\x09\x0A\x0D\x20-\x7EÀ-ÿ°º¿]/g, "")
        .trim();
    }

    function upper(value) {
      return clean(value).toUpperCase();
    }

    function formatDate(value) {
      const safe = clean(value);

      if (/^\d{4}-\d{2}-\d{2}$/.test(safe)) {
        const [yyyy, mm, dd] = safe.split("-");
        return `${dd}/${mm}/${yyyy}`;
      }

      return safe;
    }

    function findLogoPath() {
      const names = ["Cars.JPG", "cars.jpg", "Cars.jpg", "CARS.JPG", "cars.JPG"];

      for (const name of names) {
        const logoPath = path.join(__dirname, name);
        if (fs.existsSync(logoPath)) return logoPath;
      }

      return null;
    }

    function drawCenteredText(text, x, y, width, size, usedFont) {
      const safe = clean(text);
      const textWidth = usedFont.widthOfTextAtSize(safe, size);

      page.drawText(safe, {
        x: x + Math.max((width - textWidth) / 2, 2),
        y,
        size,
        font: usedFont,
        color: black
      });
    }

    function drawBox(x, y, width, height, options = {}) {
      page.drawRectangle({
        x,
        y,
        width,
        height,
        color: options.fill || undefined,
        borderWidth: options.borderWidth === undefined ? 0.7 : options.borderWidth,
        borderColor: options.borderColor || black
      });
    }

   function addField(name, value, x, y, width, height, size = 9, uppercase = false) {
      const field = form.createTextField(name);
    
      field.addToPage(page, {
        x,
        y,
        width,
        height,
        borderWidth: 0,
        textColor: black,
        backgroundColor: rgb(1, 1, 1)
      });
    
      field.setText(uppercase ? upper(value) : clean(value));
      field.setFontSize(size);
      field.defaultUpdateAppearances(font);
    
      return field;
    }



    function drawLabelField(label, name, value, x, y, width, height, labelWidth, size = 9) {
      drawBox(x, y, width, height, { borderColor: borderGray });

      page.drawText(clean(label), {
        x: x + 5,
        y: y + 7,
        size,
        font,
        color: black
      });

      addField(
        name,
        value,
        x + labelWidth,
        y + 3,
        width - labelWidth - 6,
        height - 6,
        size
      );
    }

    stage = "logo";

    const logoPath = findLogoPath();

    if (logoPath) {
      try {
        const logoBytes = fs.readFileSync(logoPath);
        const logoImage = await pdfDoc.embedJpg(logoBytes);

        page.drawImage(logoImage, {
          x: 38,
          y: 795,
          width: 110,
          height: 34
        });
      } catch (logoError) {
        console.error("ERROR INSERTANDO LOGO:", logoError);
      }
    } else {
      console.warn("No se encontró logo Cars.JPG/cars.jpg en el backend");
    }

    stage = "titulo";

    drawBox(38, 758, 519, 26, {
      fill: gray,
      borderWidth: 0
    });

    drawCenteredText("FORMULARIO EXTRA SEGURO", 38, 767, 519, 12, fontBold);

    stage = "cabecera";

    drawLabelField("Taller N°:", "taller", data.taller, 55, 718, 170, 24, 72, 9);
    drawLabelField("Serie y N°:", "serieNumero", data.serieNumero, 245, 718, 185, 24, 92, 9);
    drawLabelField("Fecha:", "fecha", formatDate(data.fecha), 450, 718, 100, 24, 38, 9);


    const leftColX = 48;  
    const canvasY = 525;
    const canvasW = 330;
    const canvasH = 165;

    stage = "canvasImage";

    if (data.canvasImage && typeof data.canvasImage === "string" && data.canvasImage.includes(",")) {
      try {
        const header = data.canvasImage.split(",")[0].toLowerCase();
        const base64 = data.canvasImage.split(",")[1];
        const imageBytes = Buffer.from(base64, "base64");

        let img = null;

        if (header.includes("image/png")) {
          img = await pdfDoc.embedPng(imageBytes);
        } else if (header.includes("image/jpeg") || header.includes("image/jpg")) {
          img = await pdfDoc.embedJpg(imageBytes);
        }

        if (img) {
          page.drawImage(img, {
            x: leftColX,
            y: canvasY,
            width: canvasW,
            height: canvasH
          });
        }
      } catch (imageError) {
        console.error("ERROR INSERTANDO canvasImage:");
        console.error(imageError);
      }
    }

    stage = "campos_derecha";

    drawBox(370, 650, 180, 24, { borderColor: borderGray });

      page.drawText("Siniestro:", {
        x: 375,
        y: 657,
        size: 9,
        font,
        color: black
      });
      
      addField("siniestro", data.siniestro1, 430, 653, 55, 14, 9);
      
      page.drawText("-", {
        x: 489,
        y: 657,
        size: 10,
        font,
        color: black
      });
      
      page.drawText("Año:", {
        x: 499,
        y: 657,
        size: 9,
        font,
        color: black
      });
      
      addField("anio", data.siniestro2, 523, 653, 26, 14, 9);



    drawLabelField("¿Dificulta visual?:", "dificultadVisual", data.dificultadVisual, 405, 606, 145, 24, 92, 9);

    stage = "texto_central";
      
      drawBox(38, 480, 519, 38, {
        borderColor: black,
        borderWidth: 0.8
      });
      
      drawCenteredText(
        "SE INFORMAN RUBROS CUYOS PORCENTAJES NO SE TENDRAN EN CUENTA EN FUTURAS RECLAMACIONES",
        42,
        503,
        511,
        8,
        fontBold
      );
      
      drawCenteredText(
        "ANULA / REMPLAZA EXTRA SEGURO DE FECHA:",
        42,
        488,
        511,
        8,
        fontBold
      );


    stage = "tablas";

    function drawTabla(tabla, prefix, x, y) {
      const rows = Array.isArray(tabla) ? tabla : [];
      const colW = [150, 50, 50];
      const tableW = colW.reduce((a, b) => a + b, 0);
      const headerH = 18;
      const rowH = 15;

      drawBox(x, y, tableW, headerH, {
        fill: gray,
        borderColor: black,
        borderWidth: 0.7
      });

      let cursorX = x;
      const headers = ["PIEZA / ACCESORIOS", "CHAPA", "PINTURA"];

      headers.forEach((header, c) => {
        drawBox(cursorX, y, colW[c], headerH, {
          fill: gray,
          borderColor: black,
          borderWidth: 0.7
        });

        drawCenteredText(header, cursorX, y + 6, colW[c], 7, fontBold);
        cursorX += colW[c];
      });

      let currentY = y - rowH;

      rows.forEach((row, i) => {
        const values = [
          row && row.pieza ? row.pieza : "",
          row && row.chapa ? row.chapa : "",
          row && row.pintura ? row.pintura : ""
        ];

        let cellX = x;

        values.forEach((value, c) => {
          drawBox(cellX, currentY, colW[c], rowH, {
            borderColor: black,
            borderWidth: 0.6
          });

          addField(
            `${prefix}_${i}_${c}`,
            value,
            cellX + 2,
            currentY + 1,
            colW[c] - 4,
            rowH - 2,
            8,
            true
          );

          cellX += colW[c];
        });

        currentY -= rowH;
      });

      return currentY;
    }

    const tableTopY = 452;

    const bottomTablesY = Math.min(
      drawTabla(data.tabla1, "tabla1", 38, tableTopY),
      drawTabla(data.tabla2, "tabla2", 307, tableTopY)
    );

    stage = "firma";

    const firmaY = bottomTablesY - 30;

    page.drawText("POR CARS:", {
      x: 42,
      y: firmaY + 5,
      size: 9,
      font: fontBold,
      color: black
    });

    addField("quien", data.quien, 95, firmaY + 1, 170, 14, 9, true);

    page.drawLine({
      start: { x: 38, y: firmaY },
      end: { x: 280, y: firmaY },
      thickness: 1.2,
      color: black
    });

    stage = "apariencias";

    form.getFields().forEach(field => {
      if (field.defaultUpdateAppearances) {
        field.defaultUpdateAppearances(font);
      }
    });
    
    form.updateFieldAppearances(font);

    stage = "guardar_pdf";

    const pdfBytes = await pdfDoc.save({
      useObjectStreams: false,
    });

    stage = "token_sharepoint";

    const token = await getAccessToken();
    const filename = `editable_${Date.now()}.pdf`;

    stage = "upload_sharepoint";

    const result = await uploadToSharePoint(
      token,
      Buffer.from(pdfBytes),
      filename,
      DEFAULT_FOLDER
    );

    res.json({
      ok: true,
      name: filename,
      webUrl: result.webUrl
    });

  } catch (err) {
    console.error("ERROR PDF EDITABLE:");
    console.error("STAGE:", stage);
    console.error(err);
    console.error(err.stack);

    res.status(500).json({
      ok: false,
      stage,
      error: err.message,
      stack: err.stack
    });
  }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log("Servidor corriendo en puerto", PORT);
});




















































































