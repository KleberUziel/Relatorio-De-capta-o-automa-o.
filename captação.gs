function onFormSubmit(e) {
  const dados = e.namedValues;
  Logger.log(JSON.stringify(dados));

  const codigoImovel = (dados["Nome do Prédio ou Residencial e unidade:"]?.[0]) || "SEM_IDENTIFICAÇÃO";
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

  const hoje = new Date();
  const proximoContato = new Date(hoje.getTime() + 15 * 24 * 60 * 60 * 1000);
  body.appendParagraph(`📅 Primeiro lembrete agendado para: ${proximoContato.toLocaleDateString("pt-BR")}`);

  doc.saveAndClose();

  const fileDoc = DriveApp.getFileById(doc.getId());
  pastaCodigo.addFile(fileDoc);
  DriveApp.getRootFolder().removeFile(fileDoc);

  const informacoesAdicionais = dados["Informações adicionais sobre o imóvel (ocupação, estado, pendências, etc):"]?.[0] || "Sem informações adicionais.";
const arquivoTxt = DriveApp.createFile(`Informações - ${codigoImovel}.txt`, informacoesAdicionais, MimeType.PLAIN_TEXT);
  pastaCodigo.addFile(arquivoTxt);
  DriveApp.getRootFolder().removeFile(arquivoTxt);

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  const ultimaLinha = planilha.getLastRow();
  let titulos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];

  const colunasExtras = ["Link da Ficha", "Link das Fotos", "Link dos Documentos"];
  const colIndices = [];

  colunasExtras.forEach((titulo) => {
    let idx = titulos.indexOf(titulo);
    if (idx === -1) {
      const novaColuna = titulos.length + 1;
      planilha.getRange(1, novaColuna).setValue(titulo);
      titulos.push(titulo);
      colIndices.push(novaColuna);
    } else {
      colIndices.push(idx + 1);
    }
  });

  planilha.getRange(ultimaLinha, colIndices[0]).setValue(fileDoc.getUrl());
  planilha.getRange(ultimaLinha, colIndices[1]).setValue(pastaFotos.getUrl());
  planilha.getRange(ultimaLinha, colIndices[2]).setValue(pastaDocs.getUrl());

  const linha = planilha.getRange(ultimaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];
  moverArquivos(linha[50], pastaFotos);
  moverArquivos(linha[51], pastaDocs);

  // ✅ Criação de lembretes no Google Agenda
  try {
    const emailCorretor = dados["Endereço de e-mail"]?.[0];
    const emailGestor = "gerencia.sunrisejp@gmail.com";

    const nomeProprietario = dados["Nome completo do proprietário:"]?.[0] || "Proprietário";
    const telefone = dados["Telefone (com DDD):"]?.[0] || "";
    const endereco = dados["Endereço completo do imóvel:"]?.[0] || "";
    const titulo = `📞 Falar com ${nomeProprietario}`;
    const descricao = `Entrar em contato com o proprietário do imóvel ${codigoImovel}.\n\nTelefone: ${telefone}\nEndereço: ${endereco}`;

    const calendario = CalendarApp.getDefaultCalendar();
    const hoje = new Date();

    for (let i = 0; i < 6; i++) {
      const dataEvento = new Date(hoje);
      dataEvento.setDate(hoje.getDate() + (15 * (i + 1)));

      calendario.createEvent(titulo, dataEvento, dataEvento, {
        description: descricao,
        guests: `${emailCorretor},${emailGestor}`,
        sendInvites: true
      });
    }
  } catch (erroAgenda) {
    Logger.log("Erro ao criar eventos no Google Agenda: " + erroAgenda);
  }
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

// 🧪 Teste manual (roda a partir da planilha, útil para debug)
function testeManual() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  const ultimaLinha = planilha.getLastRow();
  const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
  const valores = planilha.getRange(ultimaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];

  let namedValues = {};
  cabecalhos.forEach((titulo, i) => {
    namedValues[titulo] = [valores[i]];
  });

  const e = { namedValues };
  onFormSubmit(e);
}
