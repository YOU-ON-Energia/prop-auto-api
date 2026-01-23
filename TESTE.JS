function toShareId(url) {
  const b64 = Buffer.from(url, "utf8").toString("base64");
  // base64url: troca + por -, / por _, remove =
  const b64url = b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  return "u!" + b64url;
}

const url = "https://youonenergy-my.sharepoint.com/:x:/p/mzaniboni/IQBr9J5nfuy8SrzJRQ0oOataAeTsPeVDCcyRXXwK_PZThCE?e=JcZt6T";
console.log(toShareId(url));
