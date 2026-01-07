// ajustar-powerpoint.js (CommonJS) - ajusta PPTX e faz upload pro Vercel Blob
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { put } = require("@vercel/blob");

require("dotenv").config();

async function main() {
  // ====== CONFIG ======
  const inputPptx = path.join(__dirname, "templates", "atilio-com-solar.pptx"); // local
  const blobKey = "templates/atilio-com-solar.pptx"; // caminho no Blob (pasta templates)
  // Se você quiser versionar:
  // const blobKey = `templates/Proposta-YOUON-leona-${Date.now()}.pptx`;

  const token = process.env.BLOB_READ_WRITE_TOKEN;
  if (!token) {
    console.error("Faltou BLOB_READ_WRITE_TOKEN no .env");
    process.exit(1);
  }

  if (!fs.existsSync(inputPptx)) {
    console.error("Arquivo não encontrado:", inputPptx);
    process.exit(1);
  }

  // ====== LÊ + AJUSTA PPTX ======
  const buf = fs.readFileSync(inputPptx);
  const zip = await JSZip.loadAsync(buf);

  const slideFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
  );

  for (const fileName of slideFiles) {
    let xml = await zip.files[fileName].async("string");

    // EXEMPLO: aqui entram suas alterações reais
    // (coloque o que você quer substituir/ajustar de verdade)
    xml = xml.split("{NOME_CLIENTE}").join("{NOME_CLIENTE}");
    xml = xml.split("{VALOR_CONTA}").join("{VALOR_CONTA}");

    zip.file(fileName, xml);
  }

  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

  // ====== UPLOAD PRO BLOB ======
  const result = await put(blobKey, outBuffer, {
    access: "public", // pode ser public para templates
    token,            // usa seu token local
    contentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    addRandomSuffix: false,
    allowOverwrite: true,// importante: mantém o nome exato e sobrescreve
  });

  console.log("✅ Upload concluído!");
  console.log("Blob key:", blobKey);
  console.log("URL:", result.url);
}

main().catch((e) => {
  console.error("Erro:", e);
  process.exit(1);
});
