const JSZip = require("jszip");

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

      nomeCliente = "Cliente",
      inversores = "",
      baterias = "",
      energia_armazenavel = "",
      potencia = "",
      endereco = "",
      dataProposta,
      geracao = "",
      economia = "",
      unidade = "",
    } = body || {};

    // ===== escolhe template (somente 2) =====
    const solarFlag = normalizeSolarFlag(temSolar);
    const templateFile = solarFlag
    ? "YOUON_Template_Proposta_Comercial_02_ 2.pptx" // COM solar
    : "YOUON_Template_Proposta_Comercial_01_ 1.pptx"; // SEM solar


    // ===== URL do Blob (ENV) =====
    const baseUrl = process.env.TEMPLATES_BASE_URL;
    if (!baseUrl) {
      res.statusCode = 500;
      return res.end("ENV TEMPLATES_BASE_URL não configurada na Vercel.");
    }

    const templateUrl =
    `${baseUrl.replace(/\/+$/, "")}/${encodeURIComponent(templateFile)}`;


    // ===== baixa o template do Blob =====
    const response = await fetch(templateUrl);
    if (!response.ok) {
      res.statusCode = 400;
      return res.end(
        `Template não encontrado no Blob: ${templateFile}\n` +
          `URL: ${templateUrl}\n` +
          `temSolar="${temSolar}" (-> ${solarFlag})`
      );
    }

    const templateBuffer = Buffer.from(await response.arrayBuffer());

    // ===== replacements =====
    const replacements = {
      "{NOME_CLIENTE}": String(nomeCliente ?? "Cliente"),
      "{ENDERECO}": String(endereco ?? ""),
      "{DATA_PROPOSTA}": String(dataProposta || new Date().toLocaleDateString("pt-BR")),
      "{INVERSORES}": String(inversores ?? ""),
      "{POTENCIA_INVER}": String(potencia ?? ""),
      "{BATERIAS}": String(baterias ?? ""),
      "{ENERGIA_ARMAZENAVEL}": String(energia_armazenavel ?? ""),
      "{GERACAO}": String(geracao ?? ""),
      "{ECONOMIA}": String(economia ?? ""),
      "{UNIDADE}": String(unidade ?? ""),
    };

    const zip = await JSZip.loadAsync(templateBuffer);

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

    const filename = `Proposta-YOUON-${safeName}-${solarFlag ? "com-solar" : "sem-solar"}.pptx`;

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
