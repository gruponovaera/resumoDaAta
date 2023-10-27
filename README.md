# Resumo da Ata do Grupo Nova Era OnLine de NA.

Este é um projeto que utiliza o Google Apps Script para automatizar a geração de resumos a partir de uma planilha do Google Sheets e enviar esses resumos via WhatsApp.
Trata-se de um fork do projeto original do Grupo QuarenteNA e do Grupo Graça: https://github.com/grupoquarentena/resumoDaAta https://github.com/grupograca/resumoDaAta

## Como Funciona

O script lê os dados da planilha 'Respostas ao formulário 1' no Google Sheets e gera um resumo com informações relevantes. Em seguida, envia esse resumo para um grupo no WhatsApp.

## Pré-requisitos

Antes de executar este projeto, você precisará:

- Ter uma conta no Google.
- Ter acesso a uma planilha do Google Sheets.
- Utilizar a EvolutionAPI para enviar via WhatsApp.
  (https://github.com/EvolutionAPI/evolution-api)

## Configuração

1. Faça uma cópia da planilha de exemplo fornecida ou crie sua própria planilha com os dados relevantes.

2. No script `resumoDaAta.gs`, substitua as informações da API para enviar mensagens para o WhatsApp.

3. Configure um gatilho de execução para que a função `resumoDaAta` seja acionada automaticamente em intervalos específicos.

## Execução

Após configurar tudo corretamente, o script será executado automaticamente conforme configurado no gatilho. Ele lerá os dados da planilha, gerará o resumo e o enviará para o grupo no Telegram.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir problemas, propor melhorias ou enviar solicitações de pull.

## Licença

Este projeto é licenciado sob a Licença MIT.

---

**Nota:** Este é um projeto de exemplo e pode precisar de personalizações adicionais para atender às suas necessidades específicas.