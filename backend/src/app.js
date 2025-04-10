import express from "express";
import cors from 'cors';
import dotenv from "dotenv";
dotenv.config();

import emailRoutes from "./routes/emailRoutes.js";

const app = express();

app.use(express.json());

app.use(cors({
  origin: process.env.CORS_ORIGIN,
  credentials: true,
}));

app.use("/api/emails", emailRoutes);

app.get("/", (req, res) => {
  res.send("API is running...");
});

export default app;