function enviarMensagemWhatsApp(mensagem) {
  var apiUrl = "http://url.api:8080/message/sendText/instance"; // Substituir pela URL da EvolutionAPI
  var apiKey = "Token"; // Substituir pelo seu Token da API

  var payload = {
    number: "+5571999887766", // Substituir pelo número ou ID do grupo
    options: {
      delay: 1200,
      presence: "composing",
      linkPreview: false,
    },
    textMessage: {
      text: mensagem,
    },
  };

  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "apikey": apiKey,
    },
    payload: JSON.stringify(payload),
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());

  // Verifica se a mensagem foi enviada com sucesso
  if (responseData.success) {
    Logger.log("Mensagem enviada com sucesso para o WhatsApp via API.");
  } else {
    Logger.log("Erro ao enviar a mensagem para o WhatsApp via API: " + responseData.error);
  }
}

function formatarData(data) {
  var diasDaSemana = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  
  var diaSemana = diasDaSemana[data.getDay()];
  var dia = ("0" + data.getDate()).slice(-2);
  var mes = ("0" + (data.getMonth() + 1)).slice(-2);
  var ano = data.getFullYear();
  
  return diaSemana + " " + dia + "/" + mes + "/" + ano;
}

function resumoDaAta() {
  // Pegar a planilha
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pegar a aba específica das respostas
  var aba = planilha.getSheetByName('Respostas ao formulário 1');
  
  // Pegar a última linha
  var ultimaLinha = aba.getLastRow();
  
  // Definir as colunas de interesse
  var colunas = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'];
  
  // Armazenar os valores
  var valores = {};
  
  // Pegar os valores pelas colunas
  colunas.forEach(function(coluna) {
    var valor = aba.getRange(coluna + ultimaLinha).getValue();
    valores[coluna] = valor;
  });

  // Converter a data da reunião para o formato BR
  var dataReuniao = new Date(valores['A']);
  var dataFormatadaReuniao = formatarData(dataReuniao);

  // Converter a data da atualização para o formato BR
  var dataAtualiza = new Date(valores['U']);
  var dataFormatadaAtualiza = formatarData(dataAtualiza);

  
  // Montar a mensagem de resumo
  var mensagem = "*Resumo da Ata do Grupo Nova Era OnLine de NA*\n";
  
  // Linhas das mensagens necessárias
  mensagem += "*Formato da Reunião*: " + valores['B'] + "\n";
  mensagem += "*Data da Reunião*: " + dataFormatadaReuniao + "\n";
  mensagem += "*Coordenador(a)*: " + valores['C'] + "\n";
  mensagem += "*Presenças*: " + valores['D'] + "\n";
  mensagem += "*Partilhas*: " + valores['E'] + "\n";
  mensagem += "*Saldo da 7ª Tradição*: R$ " + valores['K'] + "\n";
  mensagem += "*Data da Atualiação*: " + dataFormatadaAtualiza + "\n";

  // Adicionar as linhas apenas se os valores não forem strings vazias
  if (valores['L'] !== "") {
    mensagem += "*Total de Despesas*: R$ " + valores['L'] + "\n";
  }
  if (valores['M'] !== "") {
    mensagem += "*Descrição das Despesas*: " + valores['M'] + "\n";
  }

  if (valores['F'] !== "") {
    mensagem += "*Visita(s)*: " + valores['F'] + "\n";
  }
  if (valores['G'] !== "") {
    mensagem += "*Ingresso(s)*: " + valores['G'] + "\n";
  }
  if (valores['I'] !== "") {
  mensagem += "*Nome(s) do(s) Ingressante(s)*: " + valores['I'] + "\n";
  }
  if (valores['P'] !== "") {
  mensagem += "*Contato(s) do(s) Ingressante(s)*: " + valores['P'] + "\n";
  }
  if (valores['S'] !== "") {
  mensagem += "*Visita Soube Através*: " + valores['S'] + "\n";
  }

  if (valores['H'] !== "") {
    mensagem += "*Conquista(s)*: " + valores['H'] + "\n";
  }
  if (valores['J'] !== "") {
  mensagem += "*Nome(s) da(s) Conquista(s)*: " + valores['J'] + "\n";
  }

  if (valores['N'] !== "") {
  mensagem += "*Título da Temática*: " + valores['N'] + "\n";
  }
  if (valores['O'] !== "") {
  mensagem += "*Partilhador da Temática*: " + valores['O'] + "\n";
  }

  if (valores['Q'] !== "") {
  mensagem += "*Eleição de Encargo*: " + valores['Q'] + "\n";
  }

  if (valores['R'] !== "") {
  mensagem += "*Observações*: " + valores['R'] + "\n";
  }

  if (valores['T'] !== "") {
  mensagem += "*Informações Adicionais*: " + valores['T'];
  }

  // Exibir log da mensagem montada
  Logger.log(mensagem)
  
  // Enviar a mensagem via WhatsApp (API)
  enviarMensagemWhatsApp(mensagem);
}