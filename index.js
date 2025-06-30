const express = require("express");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(express.json());

app.post("/generate", (req, res) => {
  try {
    const data = req.body; // JSON with your variables

    const content = fs.readFileSync("./template.docx", "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);

    doc.setData(data);
    doc.render();

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