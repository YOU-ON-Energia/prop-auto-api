const JSZip = require("jszip");

function normalizeSolarFlag(v) {
  if (typeof v === "boolean") return v;
  const raw = String(v || "").trim().toLowerCase();
  if (["sim", "s", "true", "1", "yes", "y"].includes(raw)) return true;
  if (["nao", "não", "n", "false", "0", "no"].includes(raw)) return false;
  return false;
}

function escapeXml(str = "") {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// substitui token mesmo quando está dividido em vários <a:t>
function replaceTokenAcrossAT(xml, token, value) {
  const reAT = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;

  const nodes = [];
  let m;
  while ((m = reAT.exec(xml)) !== null) {
    nodes.push({ start: m.index, end: reAT.lastIndex, full: m[0], text: m[1] || "" });
  }
  if (!nodes.length) return xml;

  const texts = nodes.map(n => n.text);
  let joined = texts.join("");
  let idx = joined.indexOf(token);
  if (idx === -1) return xml;

  const safeValue = escapeXml(String(value ?? ""));

  while (idx !== -1) {
    const endIdx = idx + token.length;

    let acc = 0;
    let startNode = -1, startOffset = 0;
    let endNode = -1, endOffset = 0;

    for (let i = 0; i < texts.length; i++) {
      const len = texts[i].length;
      if (startNode === -1 && acc + len > idx) { startNode = i; startOffset = idx - acc; }
      if (acc + len >= endIdx) { endNode = i; endOffset = endIdx - acc; break; }
      acc += len;
    }

    const before = texts[startNode].slice(0, startOffset);
    const after  = texts[endNode].slice(endOffset);

    texts[startNode] = before + safeValue + after;
    for (let i = startNode + 1; i <= endNode; i++) texts[i] = "";

    joined = texts.join("");
    idx = joined.indexOf(token, idx + safeValue.length);
  }

  let out = "";
  let last = 0;
  for (let i = 0; i < nodes.length; i++) {
    out += xml.slice(last, nodes[i].start);
    const openTag = nodes[i].full.match(/^<a:t[^>]*>/)?.[0] || "<a:t>";
    out += `${openTag}${texts[i] || ""}</a:t>`;
    last = nodes[i].end;
  }
  out += xml.slice(last);

  return out;
}

function applyReplacements(xml, replacements) {
  let out = xml;

  // tenta direto
  for (const [k, v] of Object.entries(replacements)) {
    out = out.split(k).join(escapeXml(String(v ?? "")));
  }

  // garante o “quebrado”
  for (const [k, v] of Object.entries(replacements)) {
    out = replaceTokenAcrossAT(out, k, v);
  }

  return out;
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

      nomeCliente = "",
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
    ? "YOUON_Template_Proposta_Comercial_02.pptx" // COM solar
    : "YOUON_Template_Proposta_Comercial_01.pptx"; // SEM solar


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
      "{NOME_CLIENTE}": String(nomeCliente ?? ""),
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

    const xmlTargets = Object.keys(zip.files).filter((name) => {
    const isRelevant =
      name.startsWith("ppt/slides/slide") ||
      name.startsWith("ppt/slideLayouts/slideLayout") ||
      name.startsWith("ppt/slideMasters/slideMaster") ||
      name.startsWith("ppt/notesSlides/notesSlide");
    return isRelevant && name.endsWith(".xml");
  });


    for (const fileName of xmlTargets) {
    const xml = await zip.files[fileName].async("string");
    const updated = applyReplacements(xml, replacements);
    zip.file(fileName, updated);
  } 


    const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

    const safeName = String(nomeCliente || "Cliente")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // remove acentos
    .replace(/[^\w\s-]/g, "")                        // remove caracteres estranhos
    .trim()
    .replace(/\s+/g, "-");                           // espaços -> hífen

    const filename = `YOUON_Template_Proposta_Comercial_${safeName}.pptx`;


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
