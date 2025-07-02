function onFormSubmit(e) {
  const dados = e.namedValues;
  Logger.log(JSON.stringify(dados));

  const codigoImovel = (dados["Nome do Prédio ou Residencial e unidade:"] && dados["Nome do Prédio ou Residencial e unidade:"][0]) || "SEM_IDENTIFICAÇÃO";
  const pastaRaiz = DriveApp.getFolderById("16EtaNrZN909h6dN2Sn0h5ER15JtBxe60");

  const pastaCodigo = pastaRaiz.createFolder(`${codigoImovel}`);
  const pastaFotos = pastaCodigo.createFolder("📸 Fotos do Imóvel");
  const pastaDocs = pastaCodigo.createFolder("📄 Documentação do Imóvel");

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

  let texto = "📋 Ficha de Captação (Resumo)\n\n";
  for (let campo in dados) {
    texto += `${campo}: ${dados[campo][0]}\n`;
  }
  const arquivoTxt = DriveApp.createFile(`Informações - ${codigoImovel}.txt`, texto, MimeType.PLAIN_TEXT);
  pastaCodigo.addFile(arquivoTxt);
  DriveApp.getRootFolder().removeFile(arquivoTxt);

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  const ultimaLinha = planilha.getLastRow();
  let titulos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];

  // Garante que os títulos existam e captura os índices
  const colunasExtras = ["Link da Ficha", "Link das Fotos", "Link dos Documentos"];
  const colIndices = [];

  colunasExtras.forEach((titulo) => {
    let idx = titulos.indexOf(titulo);
    if (idx === -1) {
      // Se não existir, cria no final
      const novaColuna = titulos.length + 1;
      planilha.getRange(1, novaColuna).setValue(titulo);
      titulos.push(titulo);
      colIndices.push(novaColuna);
    } else {
      colIndices.push(idx + 1); // índice para usar com getRange (começa em 1)
    }
  });

  // Insere os links nas colunas corretas da última linha
  planilha.getRange(ultimaLinha, colIndices[0]).setValue(fileDoc.getUrl());
  planilha.getRange(ultimaLinha, colIndices[1]).setValue(pastaFotos.getUrl());
  planilha.getRange(ultimaLinha, colIndices[2]).setValue(pastaDocs.getUrl());

  // Move arquivos enviados no formulário
  const linha = planilha.getRange(ultimaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];
  moverArquivos(linha[50], pastaFotos); // Coluna AY
  moverArquivos(linha[51], pastaDocs);  // Coluna AZ
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

// Para testar manualmente (executar diretamente)
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
