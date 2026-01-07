const JSZip = require("jszip");

// ===== helpers =====
function normalizeVendor(v) {
  const raw = String(v || "").trim().toLowerCase();

  if (raw === "matheuzinho" || raw === "mateuzinho" || raw === "mateusinho") return "mateus";
  if (raw === "atílio" || raw === "attilio") return "atilio";
  if (raw === "alex" || raw === "alê" || raw === "alê." || raw === "ale") return "ale";
  if (raw === "peixe") return "peixe";

  return raw
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeSolarFlag(v) {
  if (typeof v === "boolean") return v;
  const raw = String(v || "").trim().toLowerCase();
  if (["sim", "s", "true", "1", "yes", "y"].includes(raw)) return true;
  if (["nao", "não", "n", "false", "0", "no"].includes(raw)) return false;
  return false;
}

module.exports = async (req, res) => {
  try {
    if (req.method !== "POST") {
      res.statusCode = 405;
      return res.end("Method Not Allowed");
    }

    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;

    const {
      temSolar,
      vendedor,

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

    // ===== URL do Blob (ENV) =====
    const baseUrl = process.env.TEMPLATES_BASE_URL;
    if (!baseUrl) {
      res.statusCode = 500;
      return res.end("ENV TEMPLATES_BASE_URL não configurada na Vercel.");
    }

    const templateUrl = `${baseUrl.replace(/\/+$/, "")}/${templateFile}`;

    // ===== baixa o template do Blob =====
    const response = await fetch(templateUrl);
    if (!response.ok) {
      res.statusCode = 400;
      return res.end(
        `Template não encontrado no Blob: ${templateFile}\n` +
        `URL: ${templateUrl}\n` +
        `Verifique vendedor="${vendedor}" (-> "${vendorKey}") e temSolar="${temSolar}" (-> ${solarFlag}).`
      );
    }

    const templateBuffer = Buffer.from(await response.arrayBuffer());

    // ===== replacements =====
    const replacements = {
      "{NOME_CLIENTE}": String(nomeCliente ?? "Cliente"),
      "{VALOR_CONTA}": String(valorContaFmt ?? ""),
      "{INVERSORES}": String(inversores ?? ""),
      "{BATERIAS}": String(baterias ?? ""),
      "{ENDERECO}": String(endereco ?? ""),
      "{FAIXA}": String(faixaLabel ?? ""),
      "{DATA_PROPOSTA}": String(dataProposta || new Date().toLocaleDateString("pt-BR")),
    };

    // Abre PPTX (zip)
    const zip = await JSZip.loadAsync(templateBuffer);

    // Slides XML
    const slideFiles = Object.keys(zip.files).filter(
      (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );

    for (const fileName of slideFiles) {
      let xml = await zip.files[fileName].async("string");

      for (const [key, value] of Object.entries(replacements)) {
        xml = xml.split(key).join(value);
      }

      zip.file(fileName, xml);
    }

    const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

    const safeName = String(nomeCliente || "Cliente")
      .replace(/[^\w\s-]/g, "")
      .replace(/\s+/g, "-");

    const filename = `Proposta-YOUON-${safeName}-${vendorKey}-${suffix}.pptx`;

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
