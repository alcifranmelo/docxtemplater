const express = require("express");
const multer = require("multer");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const upload = multer({ storage: multer.memoryStorage() }); // store uploaded file in memory

const app = express();
app.use(express.json());

app.post("/generate", upload.single("template"), (req, res) => {
  try {
    const data = JSON.parse(req.body.data); // send JSON data as a text field named "data"
    const templateBuffer = req.file.buffer; // the uploaded DOCX template

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip);

    //doc.setData(data); deprecated
    doc.render(data);

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=result.docx",
    });
    res.send(buf);
  } catch (error) {
    console.error(error);
    res.status(500).send(error.message || "Error generating document");
  }
});

app.listen(3000, () => {
  console.log("Docxtemplater API listening on port 3000");
});