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
const DRIVE_ID = process.env.DRIVE_ID;

const DEFAULT_FOLDER = process.env.FOLDER_PATH || "Extra Seguro";
const allowedFolders = ["Formulario 1", "Extra Seguro", "Acta de Restos"];

const allowedOrigins = (process.env.ALLOWED_ORIGIN || "")
  .split(",")
  .filter(Boolean);

// ================= CORS =================

app.use(cors({
  origin: allowedOrigins.length > 0 ? allowedOrigins : "*",
  methods: ["POST", "OPTIONS", "GET"]
}));

app.options("*", cors());

// ================= SANITY =================

app.get("/", (req, res) => res.send("Backend funcionando"));

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

  if (!r.ok) {
    throw new Error(JSON.stringify(data));
  }

  return data.access_token;
}

// ================= SHAREPOINT UPLOAD =================

async function uploadToSharePoint(accessToken, buffer, filename, folder) {
  const safeFolder = encodeURI(folder);
  const safeName = encodeURIComponent(filename);

  const uploadUrl =
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${safeFolder}/${safeName}:/content`;

  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/pdf"
    },
    body: buffer
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text);
  }

  return response.json();
}

function textValue(value) {
  if (value === undefined || value === null) return "";
  return String(value);
}

async function embedDataUrlImage(pdfDoc, dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string" || !dataUrl.includes(",")) {
    return null;
  }

  const match = dataUrl.match(/^data:image\/(png|jpg|jpeg);base64,/i);
  if (!match) {
    return null;
  }

  const imageType = match[1].toLowerCase();
  const base64 = dataUrl.split(",")[1];

  if (!base64) {
    return null;
  }

  const bytes = Buffer.from(base64, "base64");

  if (imageType === "png") {
    return pdfDoc.embedPng(bytes);
  }

  return pdfDoc.embedJpg(bytes);
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

  } catch (error) {
    console.error("ERROR /upload:");
    console.error(error);

    res.status(500).json({
      error: error.message
    });
  }
});

// =====================================================
// ================= PDF EDITABLE ======================
// =====================================================

app.post("/generate-pdf-editable", async (req, res) => {
  try {
    const data = req.body || {};
    console.log("DATA RECIBIDA PDF EDITABLE:", data);

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
      try {
        const logoBytes = fs.readFileSync(logoPath);
        const logoImage = await pdfDoc.embedJpg(logoBytes);

        page.drawImage(logoImage, {
          x: 40,
          y: 775,
          width: 90,
          height: 35
        });
      } catch (logoError) {
        console.warn("No se pudo insertar el logo:", logoError.message);
      }
    }

    page.drawRectangle({
      x: startX,
      y: 790,
      width: totalTablesWidth,
      height: 20,
      color: rgb(0.9, 0.9, 0.9)
    });

    const titulo = "FORMULARIO EXTRA SEGURO";
    const tituloWidth = fontBold.widthOfTextAtSize(titulo, 12);

    page.drawText(titulo, {
      x: startX + (totalTablesWidth - tituloWidth) / 2,
      y: 794,
      size: 12,
      font: fontBold,
    });

    function getField(name) {
      try {
        return form.getTextField(name);
      } catch {
        return form.createTextField(name);
      }
    }

    function addTextField(name, value, x, y, width, height, fontSize = 10) {
      const field = getField(name);
      field.setText(textValue(value));
      field.setFontSize(fontSize);

      field.addToPage(page, {
        x,
        y,
        width,
        height,
        borderWidth: 0.3,
        borderColor: rgb(0.65, 0.65, 0.65),
        backgroundColor: rgb(1, 1, 1)
      });

      return field;
    }

    function drawLabelField(label, name, value, x, y, width) {
      page.drawRectangle({
        x,
        y: y - 6,
        width: width + 60,
        height: 20,
        borderWidth: 0.5,
      });

      page.drawText(label, {
        x: x + 4,
        y,
        size: 9,
        font
      });

      addTextField(name, value, x + 60, y - 5, width, 12, 10);
    }

    drawLabelField("Taller N°:", "taller", data.taller, startX, 750, 100);
    drawLabelField("Serie y N°:", "serieNumero", data.serieNumero, startX + 170, 750, 100);
    drawLabelField("Fecha:", "fecha", data.fecha, startX + 340, 750, 80);

    const leftColX = 40;
    const leftColWidth = 330;
    const rightColX = leftColX + leftColWidth + 15;
    const topY = 700;
    const imageHeight = 170;

    page.drawRectangle({
      x: leftColX,
      y: topY - imageHeight,
      width: leftColWidth,
      height: imageHeight,
      borderWidth: 0.5,
      borderColor: rgb(0.4, 0.4, 0.4)
    });

    try {
      const canvasImage = await embedDataUrlImage(pdfDoc, data.canvasImage);

      if (canvasImage) {
        page.drawImage(canvasImage, {
          x: leftColX + 1,
          y: topY - imageHeight + 1,
          width: leftColWidth - 2,
          height: imageHeight - 2
        });
      }
    } catch (imageError) {
      console.warn("No se pudo insertar la imagen del canvas:", imageError.message);
    }

    const fila1Y = topY - 20;

    page.drawText("Siniestro:", {
      x: rightColX,
      y: fila1Y,
      size: 9,
      font
    });

    addTextField(
      "siniestro",
      data.siniestro1,
      rightColX + 50,
      fila1Y - 5,
      45,
      12,
      10
    );

    page.drawText("Año:", {
      x: rightColX + 100,
      y: fila1Y,
      size: 9,
      font
    });

    addTextField(
      "anio",
      data.siniestro2,
      rightColX + 125,
      fila1Y - 5,
      35,
      12,
      10
    );

    page.drawText("Dificultad visual:", {
      x: rightColX,
      y: fila1Y - 35,
      size: 9,
      font
    });

    addTextField(
      "dificultadVisual",
      data.dificultadVisual,
      rightColX + 105,
      fila1Y - 40,
      120,
      12,
      10
    );

    function drawTabla(tabla, prefix, tableX, startY) {
      const colW = [140, 50, 50];
      const rowH = 16;
      let y = startY;

      const headers = ["Pieza", "Chapa", "Pintura"];

      headers.forEach((header, index) => {
        const x = tableX + colW.slice(0, index).reduce((a, b) => a + b, 0);

        page.drawRectangle({
          x,
          y,
          width: colW[index],
          height: rowH,
          color: rgb(0.9, 0.9, 0.9),
          borderWidth: 0.5,
          borderColor: rgb(0.4, 0.4, 0.4)
        });

        page.drawText(header, {
          x: x + 4,
          y: y + 4,
          size: 8,
          font: fontBold
        });
      });

      y -= rowH;

      (Array.isArray(tabla) ? tabla : []).forEach((row, rowIndex) => {
        const values = [
          row?.pieza || "",
          row?.chapa || "",
          row?.pintura || ""
        ];

        values.forEach((value, colIndex) => {
          const x = tableX + colW.slice(0, colIndex).reduce((a, b) => a + b, 0);

          addTextField(
            `${prefix}_${rowIndex}_${colIndex}`,
            value,
            x,
            y,
            colW[colIndex],
            rowH,
            8
          );
        });

        y -= rowH;
      });

      return y;
    }

    const tableStartY = 415;

    const bottomTablesY = Math.min(
      drawTabla(data.tabla1, "tabla1", startX, tableStartY),
      drawTabla(data.tabla2, "tabla2", startX + tableWidth + gap, tableStartY)
    );

    page.drawText("Quien realiza:", {
      x: startX,
      y: bottomTablesY - 20,
      size: 9,
      font
    });

    addTextField(
      "quien",
      data.quien,
      startX + 65,
      bottomTablesY - 24,
      140,
      14,
      10
    );

    form.updateFieldAppearances(font);

    const pdfBytes = await pdfDoc.save({
      useObjectStreams: false
    });

    const token = await getAccessToken();
    const filename = `editable_${Date.now()}.pdf`;

    const requestedFolder = data.folder?.trim();
    const folder = allowedFolders.includes(requestedFolder)
      ? requestedFolder
      : DEFAULT_FOLDER;

    const result = await uploadToSharePoint(
      token,
      Buffer.from(pdfBytes),
      filename,
      folder
    );

    res.json({
      ok: true,
      name: filename,
      webUrl: result.webUrl,
      folder
    });

  } catch (error) {
    console.error("ERROR PDF EDITABLE:");
    console.error(error);
    console.error(error.stack);

    res.status(500).json({
      error: error.message,
      stack: error.stack
    });
  }
});

// ================= START SERVER =================

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log("Servidor corriendo en puerto", PORT);
});
















































































