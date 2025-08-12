require('dotenv').config();

const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const puppeteer = require("puppeteer");
const pLimit = require("p-limit").default;

const app = express();
const port = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static("public"));

const upload = multer({ dest: "uploads/" });

const AAT_HEADER_GOL = process.env.AAT_HEADER_GOL;
const AZUL_KEY = process.env.AZUL_KEY;

let progressoAtual = 0;
let progressoTotal = 0;

app.get("/", (req, res) => res.render("index", { results: null }));

app.get("/progresso", (req, res) => {
  res.json({
    atual: progressoAtual,
    total: progressoTotal,
    porcentagem: progressoTotal ? Math.round((progressoAtual / progressoTotal) * 100) : 0
  });
});

app.post("/upload", upload.single("planilha"), async (req, res) => {
  try {
    const ext = path.extname(req.file.originalname).toLowerCase();
    let dataRows = [];

    if (ext === '.csv') {
      const fileContent = fs.readFileSync(req.file.path, 'latin1');
      const csvString = fileContent.replace(/;/g, ',');
      const workbook = xlsx.read(csvString, { type: 'string' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      dataRows = xlsx.utils.sheet_to_json(sheet, { defval: '' });
    } else if (ext === '.xlsx' || ext === '.xls') {
      const workbook = xlsx.readFile(req.file.path);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      dataRows = xlsx.utils.sheet_to_json(sheet, { defval: '' });
    } else {
      throw new Error('Tipo de arquivo não suportado. Utilize CSV ou Excel.');
    }

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
    }

    let tokenAzul = null;
    if (AZUL_KEY) {
      try {
        const tokenResp = await axios.post("https://b2c-api.voeazul.com.br/authentication/api/authentication/v1/token", {}, {
          headers: {
            "Ocp-Apim-Subscription-Key": AZUL_KEY,
            "User-Agent": "PostmanRuntime/7.36.3",
            "Accept": "*/*",
            "Content-Type": "application/json",
          }
        });
        tokenAzul = tokenResp.data?.data;
      } catch (err) {
        console.error("Erro ao obter token AZUL:", err.message);
      }
    }

    progressoAtual = 0;
    progressoTotal = dataRows.length;

    const limit = pLimit(5);

    const tasks = dataRows.map(row => limit(async () => {
      const rawCompanhia = (row["Companhia"] || '').toLowerCase();
      const isGol = rawCompanhia.includes("gol");
      const isAzul = rawCompanhia.includes("azul");
      const isLatam = rawCompanhia.includes("latam");

      const nome = (row["Nome"] || '').trim();
      const sobrenome = nome.split(" ").pop().toUpperCase();
      const origemSigla = (row["Origem"] || '').match(/\((\w{3})\)/)?.[1] || '';
      const loc = (row["Localizador"] || '').toString().trim();
      const dataPlanilha = (row["Data Embarque"] || '').toString().trim();

      try {
        if (isGol && tokenGol) {
          const url = `https://booking-api.voegol.com.br/api/pnrBnpl/pnr-bnpl-validation?context=b2c&flow=consult&pnr=${loc}&origin=${origemSigla}&lastName=${sobrenome}`;
          const resGol = await axios.get(url, { headers: { Authorization: `Bearer ${tokenGol}` } });
          const pnr = resGol.data.response?.pnrRetrieveResponse?.pnr || resGol.data;
          const parts = pnr.itinerary?.itineraryParts || [];
          const altered = parts.some(part =>
            part.segments?.some(s => s.origin === origemSigla && ["CANCELLED", "SCHEDULE_CHANGE"].includes(s.segmentStatusCode?.segmentStatus))
          );
          const dataVoo = parts[0]?.segments?.[0]?.departureDateTime || dataPlanilha;
          progressoAtual++;
          return {
            nome, localizador: loc, origem: origemSigla, sobrenome, companhia: 'Gol',
            status: altered ? 'Alterado' : 'OK',
            data: dataVoo.split("T")[0] || dataPlanilha
          };
        } else if (isAzul && tokenAzul) {
          const consultaAzul = await axios.post(
            `https://b2c-api.voeazul.com.br/canonical/api/booking/v5/bookings/${loc}`,
            { departureStation: origemSigla },
            {
              headers: {
                Authorization: `Bearer ${tokenAzul}`,
                Device: 'novosite',
                "Ocp-Apim-Subscription-Key": AZUL_KEY,
                "User-Agent": "PostmanRuntime/7.36.3",
                "Accept": "*/*",
                "Content-Type": "application/json",
              }
            }
          );
          const altered = Boolean(consultaAzul.data?.data?.journeys?.[0]?.reaccommodation?.reaccommodate);
          const dataVoo = consultaAzul.data?.data?.journeys?.[0]?.flights?.[0]?.departureDate || dataPlanilha;
          progressoAtual++;
          return {
            nome, localizador: loc, origem: origemSigla, sobrenome, companhia: 'Azul',
            status: altered ? 'Alterado' : 'OK',
            data: dataVoo.split("T")[0] || dataPlanilha
          };
        } else if (isLatam) {
          const pedido = (row["Nº da Compra"] || '').toString().trim();
          if (!pedido) throw new Error("Pedido não fornecido");

          const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
          const page = await browser.newPage();
          const url = `https://www.latamairlines.com/br/pt/minhas-viagens/second-detail?orderId=${encodeURIComponent(pedido)}&lastname=${encodeURIComponent(sobrenome)}`;
          await page.goto(url, { waitUntil: ['domcontentloaded', 'networkidle2'], timeout: 60000 });

          let altered = false;
          let dataVoo = dataPlanilha;
          try {
            await page.waitForSelector('[data-testid="status-icon-warning"]', { timeout: 10000 });
            altered = true;
          } catch {}

          try {
            dataVoo = await page.$eval('[data-testid="itinerary-date"]', el => el.textContent.trim());
          } catch {}

          await browser.close();
          progressoAtual++;
          return {
            nome, localizador: pedido, origem: origemSigla, sobrenome, companhia: 'Latam Airlines',
            status: altered ? 'Alterado' : 'OK',
            data: dataVoo
          };
        } else {
          progressoAtual++;
          return {
            nome, localizador: loc, origem: origemSigla, sobrenome, companhia: row["Companhia"] || '',
            status: 'Não compatível',
            data: dataPlanilha
          };
        }
      } catch (err) {
        progressoAtual++;
        return {
          nome, localizador: loc, origem: origemSigla, sobrenome, companhia: row["Companhia"] || '',
          status: 'Erro',
          data: dataPlanilha
        };
      }
    }));

    const results = await Promise.allSettled(tasks);
    const formattedResults = results.map(r => r.value || { status: 'Erro', data: 'Indefinido' });

    fs.unlinkSync(req.file.path);
    res.render("index", { results: formattedResults });

  } catch (err) {
    console.error(err);
    res.status(500).send(err.message || 'Erro no processamento');
  }
});

app.listen(port, () => console.log(`Servidor rodando em http://localhost:${port}`));
