import express from "express";
import cors from "cors";
import { router as convertRouter } from "./routes/convert.js";

const app = express();
const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;

app.use(cors());
app.use(express.json({ limit: "100mb" }));

app.use("/api", convertRouter);

app.get("/health", (_req, res) => res.json({ status: "ok" }));

app.listen(PORT, () => {
  console.log(`PDF to Office server running on http://localhost:${PORT}`);
});
