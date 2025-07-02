function onFormSubmit(e) {
  const dados = e.namedValues;
  Logger.log(JSON.stringify(dados)); // debug

  const codigoImovel = (dados["Código do Imóvel"] && dados["Código do Imóvel"][0]) || "SEM_CODIGO";
  const pastaRaiz = DriveApp.getFolderById("16EtaNrZN909h6dN2Sn0h5ER15JtBxe60");

  const pastaCodigo = pastaRaiz.createFolder(`${codigoImovel}`);
  const pastaFotos = pastaCodigo.createFolder("📸 Fotos do Imóvel");
  const pastaDocs = pastaCodigo.createFolder("📄 Documentação do Imóvel");

  // Cria o Google Docs com as informações do formulário
  const doc = DocumentApp.create(`Ficha de Captação - ${codigoImovel}`);
  const body = doc.getBody();
  body.appendParagraph("📋 Ficha de Captação").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  for (let campo in dados) {
    body.appendParagraph(`${campo}: ${dados[campo][0]}`);
  }
  doc.saveAndClose();

  const fileDoc = DriveApp.getFileById(doc.getId());
  pastaCodigo.addFile(fileDoc);
  DriveApp.getRootFolder().removeFile(fileDoc);

  // Cria um arquivo .txt com resumo
  let texto = "📋 Ficha de Captação (Resumo)\n\n";
  for (let campo in dados) {
    texto += `${campo}: ${dados[campo][0]}\n`;
  }
  const arquivoTxt = DriveApp.createFile(`Informações - ${codigoImovel}.txt`, texto, MimeType.PLAIN_TEXT);
  pastaCodigo.addFile(arquivoTxt);
  DriveApp.getRootFolder().removeFile(arquivoTxt);

  // Localiza planilha e pega arquivos enviados
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  const ultimaLinha = planilha.getLastRow();
  const linha = planilha.getRange(ultimaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];

  // Move arquivos enviados (colunas AY e AZ = índices 50 e 51)
  moverArquivos(linha[50], pastaFotos); // AY
  moverArquivos(linha[51], pastaDocs);  // AZ

  // ADICIONA LINKS AO FINAL DA PLANILHA SEM ALTERAR CABEÇALHOS
  const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];

  // Adiciona colunas se ainda não existirem
  if (!cabecalhos.includes("Link da Ficha")) {
    planilha.getRange(1, planilha.getLastColumn() + 1).setValue("Link da Ficha");
    planilha.getRange(1, planilha.getLastColumn() + 1).setValue("Link das Fotos");
    planilha.getRange(1, planilha.getLastColumn() + 1).setValue("Link da Documentação");
  }

  // Pega os índices atualizados
  const novosCabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
  const colFicha = novosCabecalhos.indexOf("Link da Ficha") + 1;
  const colFotos = novosCabecalhos.indexOf("Link das Fotos") + 1;
  const colDocs = novosCabecalhos.indexOf("Link da Documentação") + 1;

  // Preenche os links na linha correta
  if (colFicha) planilha.getRange(ultimaLinha, colFicha).setValue(fileDoc.getUrl());
  if (colFotos) planilha.getRange(ultimaLinha, colFotos).setValue(pastaFotos.getUrl());
  if (colDocs)  planilha.getRange(ultimaLinha, colDocs).setValue(pastaDocs.getUrl());
}

function moverArquivos(urls, destino) {
  if (!urls) return;
  const links = urls.split(', ');
  links.forEach(link => {
    const id = link.match(/[-\w]{25,}/);
    if (id && id[0]) {
      try {
        const file = DriveApp.getFileById(id[0]);
        file.moveTo(destino);
      } catch (err) {
        Logger.log("Erro ao mover arquivo: " + err);
      }
    }
  });
}

function testeManual() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  const ultimaLinha = sheet.getLastRow();
  const cabecalhos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valores = sheet.getRange(ultimaLinha, 1, 1, sheet.getLastColumn()).getValues()[0];

  let namedValues = {};
  cabecalhos.forEach((titulo, i) => {
    namedValues[titulo] = [valores[i]];
  });

  const e = { namedValues };
  onFormSubmit(e);
}
