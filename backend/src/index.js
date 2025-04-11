import dotenv from "dotenv";
dotenv.config();

import connectDB from "./db/index.js";
import app from './app.js';

connectDB()
  .then(() => {
    app.listen(process.env.PORT || 5005, () => {
      console.log(` Server is running on port: ${process.env.PORT}`);
    });
  })
  .catch((err) => {
    console.log(" MDB connection failed:", err);
  });