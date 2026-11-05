import { Router } from "express";
import multer from "multer";
import { extractFromPdf } from "../services/pdfExtractor.js";
import { tablesToExcelBuffer } from "../services/excelExporter.js";
import { tablesToWordBuffer } from "../services/wordExporter.js";
import { layoutToExcelBuffer } from "../services/excelLayoutExporter.js";
import { layoutToWordBuffer } from "../services/wordLayoutExporter.js";
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
export const router = Router();
router.post("/extract", upload.single("file"), async (req, res) => {
    try {
        if (!req.file)
            return res.status(400).json({ error: "No PDF file uploaded (field name: file)" });
        if (req.file.mimetype !== "application/pdf") {
            return res.status(400).json({ error: "Only PDF files are supported" });
        }
        const result = await extractFromPdf(req.file.buffer, req.file.originalname);
        res.json(result);
    }
    catch (err) {
        console.error("Extraction failed:", err);
        res.status(500).json({ error: "Failed to extract tables from PDF" });
    }
});
router.post("/export/excel", async (req, res) => {
    try {
        const tables = req.body.tables;
        if (!Array.isArray(tables))
            return res.status(400).json({ error: "Body must include a tables array" });
        const buffer = await tablesToExcelBuffer(tables);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", 'attachment; filename="extracted-tables.xlsx"');
        res.send(buffer);
    }
    catch (err) {
        console.error("Excel export failed:", err);
        res.status(500).json({ error: "Failed to build Excel file" });
    }
});
router.post("/export/word", async (req, res) => {
    try {
        const { tables, fileName } = req.body;
        if (!Array.isArray(tables))
            return res.status(400).json({ error: "Body must include a tables array" });
        const buffer = await tablesToWordBuffer(tables, fileName ?? "Extracted Tables");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.setHeader("Content-Disposition", 'attachment; filename="extracted-tables.docx"');
        res.send(buffer);
    }
    catch (err) {
        console.error("Word export failed:", err);
        res.status(500).json({ error: "Failed to build Word file" });
    }
});
// Layout-preserving exports — cells placed at their original PDF positions
router.post("/export/excel-layout", async (req, res) => {
    try {
        const cells = req.body.cells;
        if (!Array.isArray(cells))
            return res.status(400).json({ error: "Body must include a cells array" });
        const buffer = await layoutToExcelBuffer(cells);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", 'attachment; filename="extracted-layout.xlsx"');
        res.send(buffer);
    }
    catch (err) {
        console.error("Excel layout export failed:", err);
        res.status(500).json({ error: "Failed to build Excel layout file" });
    }
});
router.post("/export/word-layout", async (req, res) => {
    try {
        const { cells, fileName } = req.body;
        if (!Array.isArray(cells))
            return res.status(400).json({ error: "Body must include a cells array" });
        const buffer = await layoutToWordBuffer(cells, fileName ?? "Extracted Layout");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.setHeader("Content-Disposition", 'attachment; filename="extracted-layout.docx"');
        res.send(buffer);
    }
    catch (err) {
        console.error("Word layout export failed:", err);
        res.status(500).json({ error: "Failed to build Word layout file" });
    }
});
