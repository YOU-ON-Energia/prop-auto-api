const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

// ===== helpers =====
function normalizeVendor(v) {
  const raw = String(v || "").trim().toLowerCase();

  // ajuste apelidos -> prefixo do arquivo
  // (de acordo com seus templates: ale, atilio, mateus, peixe)
  if (raw === "matheuzinho" || raw === "mateuzinho" || raw === "mateusinho") return "mateus";
  if (raw === "atílio" || raw === "attilio") return "atilio";
  if (raw === "alex" || raw === "alê" || raw === "alê." || raw === "ale") return "ale";
  if (raw === "peixe") return "peixe";

  // fallback: normaliza e usa como prefixo
  return raw
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // tira acentos
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeSolarFlag(v) {
  // aceita boolean, "Sim/Não", "true/false", "1/0"
  if (typeof v === "boolean") return v;
  const raw = String(v || "").trim().toLowerCase();
  if (["sim", "s", "true", "1", "yes", "y"].includes(raw)) return true;
  if (["nao", "não", "n", "false", "0", "no"].includes(raw)) return false;
  return false;
}

module.exports = async (req, res) => {
  try {
    // Só aceita POST
    if (req.method !== "POST") {
      res.statusCode = 405;
      return res.end("Method Not Allowed");
    }

    // Vercel Serverless: req.body pode vir como string
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;

    const {
      // ==== novos campos para escolher template ====
      temSolar,      // ex: true/false ou "Sim"/"Não"
      vendedor,      // ex: "Matheuzinho", "Atilio", "Ale", "Peixe"

      // ==== seus campos atuais ====
      nomeCliente = "Cliente",
      valorContaFmt = "",
      inversores = "",
      baterias = "",
      faixaLabel = "",
      endereco = "",
      dataProposta,
    } = body || {};

    // ===== escolhe template =====
    const vendorKey = normalizeVendor(vendedor);
    const solarFlag = normalizeSolarFlag(temSolar);
    const suffix = solarFlag ? "com-solar" : "sem-solar";

    const templateFile = `${vendorKey}-${suffix}.pptx`;
    const templatePath = path.join(process.cwd(), "templates", templateFile);

    if (!fs.existsSync(templatePath)) {
      res.statusCode = 400;
      return res.end(
        `Template não encontrado: /templates/${templateFile}\n` +
        `Verifique vendedor="${vendedor}" (-> "${vendorKey}") e temSolar="${temSolar}" (-> ${solarFlag}).`
      );
    }

    // ===== replacements (igual o seu) =====
    const replacements = {
      "{NOME_CLIENTE}": String(nomeCliente ?? "Cliente"),
      "{VALOR_CONTA}": String(valorContaFmt ?? ""),
      "{INVERSORES}": String(inversores ?? ""),
      "{BATERIAS}": String(baterias ?? ""),
      "{ENDERECO}": String(endereco ?? ""),
      "{FAIXA}": String(faixaLabel ?? ""),
      "{DATA_PROPOSTA}": String(dataProposta || new Date().toLocaleDateString("pt-BR")),
      // opcional, se você quiser usar no PPT:
      // "{COM_OU_SEM_SOLAR}": solarFlag ? "COM SOLAR" : "SEM SOLAR",
    };

    // Lê o template do disco
    const templateBuffer = fs.readFileSync(templatePath);

    // Abre PPTX (zip)
    const zip = await JSZip.loadAsync(templateBuffer);

    // Slides XML
    const slideFiles = Object.keys(zip.files).filter(
      (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );

    for (const fileName of slideFiles) {
      let xml = await zip.files[fileName].async("string");

      // Replaces
      for (const [key, value] of Object.entries(replacements)) {
        xml = xml.split(key).join(value);
      }

      zip.file(fileName, xml);
    }

    // Gera PPTX final
    const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

    // Nome do arquivo
    const safeName = String(nomeCliente || "Cliente")
      .replace(/[^\w\s-]/g, "")
      .replace(/\s+/g, "-");

    const filename = `Proposta-YOUON-${safeName}-${vendorKey}-${suffix}.pptx`;

    // Headers
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.statusCode = 200;
    return res.end(outBuffer);
  } catch (err) {
    res.statusCode = 500;
    return res.end(`Erro ao gerar PPT: ${err?.message || err}`);
  }
};
