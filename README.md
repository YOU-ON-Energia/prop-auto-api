## API PARA GERAR PROPOSTA ORIENTATIVA AUTOMÁTICA (INSTRUÇÕES)
# API/api-com-sem-solar
- Api responsável por fazer a lógica se o cliente tem solar ou não, utilizando o formulário do n8n

# API/generate-ppt
-  Gera um powerpoint a partir das infos que o cliente responder no formulário da shopify

# /ajustar-powerpoint

- Responsável por enviar um powerpoint editado ou um novo powerpoint que esta na pasta templates, para o blob da vercel (banco de dados com os arquivos)

` node ajustar-powerpoint.js `
- comando responsável por enviar esse arquivo no blob

obs: necessário mudar a const 
`inputPptx` e a const `blobKey` com o nome exato do pptx que esta no templates