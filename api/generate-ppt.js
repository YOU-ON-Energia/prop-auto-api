const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

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
      nomeCliente = "Cliente",
      valorContaFmt = "",
      inversores = "",
      baterias = "",
      faixaLabel = "",
      endereco = "",
      dataProposta, 
    } = body || {};

    const replacements = {
      "{NOME_CLIENTE}": String(nomeCliente ?? "Cliente"),
      "{VALOR_CONTA}": String(valorContaFmt ?? ""),
      "{INVERSORES}": String(inversores ?? ""),
      "{BATERIAS}": String(baterias ?? ""),
      "{ENDERECO}": String(endereco ?? ""),
      "{FAIXA}": String(faixaLabel ?? ""),
      "{DATA_PROPOSTA}": String(
        dataProposta || new Date().toLocaleDateString("pt-BR")
      ),
    };

    // Lê o template do disco
    const templatePath = path.join(process.cwd(), "templates", "template.pptx");
    if (!fs.existsSync(templatePath)) {
      res.statusCode = 500;
      return res.end("Template não encontrado em /templates/template.pptx");
    }

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
        // split/join é bem compatível
        xml = xml.split(key).join(value);
      }

      zip.file(fileName, xml);
    }

    // Gera PPTX final
    const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

    // Nome do arquivo
    const safeName = String(nomeCliente || "Cliente").replace(/[^\w\s-]/g, "").replace(/\s+/g, "-");
    const filename = `Proposta-YOUON-${safeName}.pptx`;

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
