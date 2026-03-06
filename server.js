const express = require("express");
const multer = require("multer");
const cors = require("cors");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer();

app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended:true, limit:"20mb"}));

app.use(cors());

app.get("/",(req,res)=>{
  res.send("✅ Backend funcionando");
});

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


// =================================================
// GENERAR PDF EDITABLE
// =================================================

app.post("/generate-pdf-editable", async (req,res)=>{

try{

const data = req.body;

console.log("DATA RECIBIDA:",data);

const pdfDoc = await PDFDocument.create();
const page = pdfDoc.addPage([595,842]);
const form = pdfDoc.getForm();

const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);


// =================================================
// VARIABLES
// =================================================

const tableWidth = 240;
const gap = 20;

const totalTablesWidth = tableWidth*2 + gap;

const pageWidth = 595;

const startX = (pageWidth - totalTablesWidth)/2;


// =================================================
// LOGO
// =================================================

const logoPath = path.join(__dirname,"cars.jpg");

if(fs.existsSync(logoPath)){

const logoBytes = fs.readFileSync(logoPath);

const logoImage = await pdfDoc.embedJpg(logoBytes);

page.drawImage(logoImage,{
x:startX,
y:775,
width:90,
height:35
});

}


// =================================================
// TITULO
// =================================================

page.drawRectangle({
x:startX,
y:790,
width:totalTablesWidth,
height:20,
color:rgb(0.9,0.9,0.9)
});

const titulo="FORMULARIO EXTRA SEGURO";
const tituloSize=12;

const tituloWidth = fontBold.widthOfTextAtSize(titulo,tituloSize);

const tituloX = startX + (totalTablesWidth - tituloWidth)/2;

page.drawText(titulo,{
x:tituloX,
y:794,
size:tituloSize,
font:fontBold
});


// =================================================
// FUNCION CAMPOS SUPERIORES
// =================================================

function drawLabelField(label,name,x,y,width){

const boxWidth = width + 60;
const boxHeight = 20;

page.drawRectangle({
x:x,
y:y-6,
width:boxWidth,
height:boxHeight,
borderWidth:0.5,
borderColor:rgb(0,0,0)
});

page.drawText(label,{
x:x+4,
y:y,
size:9,
font
});

const f = form.createTextField(name);

f.setText(data[name] || "");

f.setTextColor(rgb(0,0,0));

f.addToPage(page,{
x:x+60,
y:y-5,
width:width,
height:12
});

}

drawLabelField("Taller N°:","taller",startX,750,100);
drawLabelField("Serie y N°:","serieNumero",210,750,100);
drawLabelField("Fecha:","fecha",380,750,80);


// =================================================
// BLOQUE IMAGEN
// =================================================

const leftColX = startX;
const leftColWidth = 330;

const rightColX = leftColX + leftColWidth + 15;

const topY = 700;
const imageHeight = 170;

if(data.canvasImage){

const img = await pdfDoc.embedPng(
Buffer.from(data.canvasImage.split(",")[1],"base64")
);

page.drawImage(img,{
x:leftColX,
y:topY-imageHeight,
width:leftColWidth,
height:imageHeight
});

page.drawRectangle({
x:leftColX,
y:topY-imageHeight,
width:leftColWidth,
height:imageHeight,
borderWidth:0.5
});

}


// =================================================
// SINIESTRO + AÑO
// =================================================

const fila1Y = topY - 20;

page.drawRectangle({
x:rightColX,
y:fila1Y-8,
width:170,
height:20,
borderWidth:0.5
});

page.drawText("Siniestro:",{
x:rightColX+4,
y:fila1Y,
size:9,
font
});

const siniestroField = form.createTextField("siniestro");

siniestroField.setText(data.siniestro1 || "");

siniestroField.addToPage(page,{
x:rightColX+50,
y:fila1Y-5,
width:45,
height:12
});

page.drawText("Año:",{
x:rightColX+100,
y:fila1Y,
size:9,
font
});

const anioField = form.createTextField("anio");

anioField.setText(data.siniestro2 || "");

anioField.addToPage(page,{
x:rightColX+125,
y:fila1Y-5,
width:35,
height:12
});


// =================================================
// DIFICULTAD VISUAL
// =================================================

const dificultadY = fila1Y - 35;

const dificultadWidth = startX + totalTablesWidth - rightColX;

page.drawRectangle({
x:rightColX,
y:dificultadY-8,
width:dificultadWidth,
height:20,
borderWidth:0.5
});

page.drawText("¿Dificultad visual?:",{
x:rightColX+4,
y:dificultadY,
size:9,
font
});

const dificultadVisualField = form.createTextField("dificultadVisual");

dificultadVisualField.setText(data.dificultadVisual || "");

dificultadVisualField.addToPage(page,{
x:rightColX+105,
y:dificultadY-5,
width:dificultadWidth-110,
height:12
});


// =================================================
// TABLAS
// =================================================

function drawTabla(tabla,startX,startY){

const colW=[140,50,50];
const headers=["PIEZA","CHAPA","PINTURA"];
const rowH=16;

let y=startY;

headers.forEach((h,i)=>{

const x=startX+colW.slice(0,i).reduce((a,b)=>a+b,0);

page.drawRectangle({
x,
y,
width:colW[i],
height:rowH,
color:rgb(0.9,0.9,0.9),
borderWidth:0.5
});

page.drawText(h,{
x:x+4,
y:y+4,
size:8,
font:fontBold
});

});

y-=rowH;

(tabla||[]).forEach((row,i)=>{

const vals=[row.pieza,row.chapa,row.pintura];

vals.forEach((val,c)=>{

const x=startX+colW.slice(0,c).reduce((a,b)=>a+b,0);

const f=form.createTextField(`tbl_${startX}_${i}_${c}`);

f.setText(val||"");

if(c===0){
f.setAlignment(0);
}

f.addToPage(page,{
x,
y,
width:colW[c],
height:rowH
});

page.drawRectangle({
x,
y,
width:colW[c],
height:rowH,
borderWidth:0.5
});

});

y-=rowH;

});

return y;

}

const tableStartY = 415;

const endY1 = drawTabla(data.tabla1,startX,tableStartY);
const endY2 = drawTabla(data.tabla2,startX+tableWidth+gap,tableStartY);

const bottomTablesY = Math.min(endY1,endY2);


// =================================================
// POR CARS
// =================================================

const quienY = bottomTablesY - 20;

page.drawText("POR CARS:",{
x:startX,
y:quienY,
size:9,
font
});

const quienField = form.createTextField("quien");

quienField.setText(data.quien || "");

quienField.addToPage(page,{
x:startX+65,
y:quienY-4,
width:140,
height:14
});

page.drawRectangle({
x:startX+65,
y:quienY-4,
width:140,
height:14,
borderWidth:0.5
});


// =================================================
// GUARDAR
// =================================================

form.updateFieldAppearances(font);

const pdfBytes = await pdfDoc.save();

res.setHeader("Content-Type","application/pdf");
res.setHeader(
"Content-Disposition",
`attachment; filename="editable_${Date.now()}.pdf"`
);

res.send(Buffer.from(pdfBytes));

}catch(err){

console.error(err);
res.status(500).json({error:err.message});

}

});


// =================================================
// START SERVER
// =================================================

const PORT = process.env.PORT || 10000;

app.listen(PORT,()=>{
console.log("Servidor corriendo en puerto",PORT);
});








































































