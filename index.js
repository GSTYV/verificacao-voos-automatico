require('dotenv').config();

const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const puppeteer = require("puppeteer");

const app = express();
const port = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static("public"));

const upload = multer({ dest: "uploads/" });

// Chaves carregadas do arquivo .env
const AAT_HEADER_GOL = process.env.AAT_HEADER_GOL;
const AZUL_KEY = process.env.AZUL_KEY;

app.get("/", (req, res) => {
  res.render("index", { results: null });
});

app.post("/upload", upload.single("planilha"), async (req, res) => {
  try {
    const workbook = xlsx.readFile(req.file.path, { FS: ";" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { defval: "", raw: false });

    let tokenGol = null;
    if (AAT_HEADER_GOL) {
      try {
        const tokenRes = await axios.get(
          "https://gol-auth-api.voegol.com.br/api/authentication/create-token",
          { headers: { "x-aat": AAT_HEADER_GOL } }
        );
        tokenGol = tokenRes.data.response.token;
      } catch (err) {
        console.error("Erro ao obter token GOL:", err.message);
      }
    } else {
      console.warn("AAT_HEADER_GOL não definido no .env");
    }

    const results = [];

    for (const row of data) {
      const companhia = row["Companhia"]?.toLowerCase();
      const localizador = row["Localizador"]?.trim();
      const origemSigla = row["Origem"]?.match(/\((\w{3})\)/)?.[1];
      const nome = row["Nome"];
      const ultimoSobrenome = nome?.trim().split(" ").pop().toUpperCase();

      if (companhia === "gol" && tokenGol) {
        try {
          const url =
            `https://booking-api.voegol.com.br/api/pnrBnpl/pnr-bnpl-validation?context=b2c&flow=consult&pnr=${localizador}&origin=${origemSigla}&lastName=${ultimoSobrenome}`;
          const consulta = await axios.get(url, {
            headers: { Authorization: `Bearer ${tokenGol}` },
          });

          const pnrData = consulta.data.response?.pnrRetrieveResponse?.pnr || consulta.data;
          const itineraryParts = pnrData.itinerary?.itineraryParts || [];

          let isAltered = false;
          for (const part of itineraryParts) {
            for (const seg of part.segments || []) {
              if (seg.origin === origemSigla) {
                const st = seg.segmentStatusCode?.segmentStatus;
                if (st === "CANCELLED" || st === "SCHEDULE_CHANGE") {
                  isAltered = true;
                  break;
                }
              }
            }
            if (isAltered) break;
            if (
              (part.cancelledSegments || []).some(
                (cs) => cs.origin === origemSigla && cs.segmentStatusCode?.segmentStatus === "CANCELLED"
              )
            ) {
              isAltered = true;
              break;
            }
          }

          results.push({
            nome,
            localizador,
            origem: origemSigla,
            sobrenome: ultimoSobrenome,
            companhia: "Gol",
            status: isAltered ? "Alterado" : "OK",
          });
        } catch (err) {
          console.error("Erro ao consultar Gol:", err.message);
          results.push({ nome, localizador, origem: origemSigla, sobrenome: ultimoSobrenome, companhia: "Gol", status: "Erro" });
        }
      } else if (companhia === "azul") {
        if (!AZUL_KEY) console.warn("AZUL_KEY não definido no .env");
        try {
          const tokenRes = await axios.post(
            "https://b2c-api.voeazul.com.br/authentication/api/authentication/v1/token",
            {},
            { headers: { "Ocp-Apim-Subscription-Key": AZUL_KEY } }
          );

          const tokenAzul = tokenRes.data?.data;
          const consultaRes = await axios.post(
            `https://b2c-api.voeazul.com.br/canonical/api/booking/v5/bookings/${localizador}`,
            { departureStation: origemSigla },
            { headers: { Authorization: `Bearer ${tokenAzul}`, "Ocp-Apim-Subscription-Key": AZUL_KEY } }
          );

          const alterado = !!consultaRes.data?.data?.reaccommodation?.redirectLinks;
          results.push({ nome, localizador, origem: origemSigla, sobrenome: ultimoSobrenome, companhia: "Azul", status: alterado ? "Alterado" : "OK" });
        } catch (err) {
          console.error("Erro ao consultar Azul:", err.message);
          results.push({ nome, localizador, origem: origemSigla, sobrenome: ultimoSobrenome, companhia: "Azul", status: "Erro" });
        }
      }
    }

    fs.unlinkSync(req.file.path);
    res.render("index", { results });
  } catch (error) {
    console.error(error);
    res.status(500).send("Erro no processamento da planilha.");
  }
});

app.listen(port, () => console.log(`Servidor rodando em http://localhost:${port}`));