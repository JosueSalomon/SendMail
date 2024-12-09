import express from "express";
import { transporter } from "./nodemailer.js";

const app = express();

app.listen(9000, () => {
  console.log("Server listen on port 9000");
});


