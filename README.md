# PROCESSOS_GEP_Automacao

Pipeline em PowerShell para automatizar o tratamento de processos de peritagem GEP.

## Estrutura

```
H:\PROCESSOS_GEP_Automacao\
??? processar_processo.ps1      # Script principal (um clique)
??? _config\
?   ??? config.ini              # Caminhos configuraveis (pdftotext, tesseract)
??? templates\
?   ??? email_template.html     # Modelo de e-mail (HTML)
??? README.md                   # Este guia
```

Ao executar o script, a estrutura abaixo e criada automaticamente para cada processo:

```
H:\PROCESSOS_GEP\<processo_id>\
??? origem\    # Ficheiro PDF original e DOCX opcional
??? trabalho\  # doc.txt, data.json, relatorio_fillform.pdf, log.txt
??? output\    # Reservado para entregaveis futuros
```

## Configuracao

Os caminhos dos executaveis externos estao centralizados em `_config\config.ini`:

```
[paths]
pdftotext = H:\Programas instalados\Ferramentas\Poppler\bin\pdftotext.exe
tesseract = H:\Programas instalados\Ferramentas\Tesseract\tesseract.exe
```

Edite esse ficheiro para apontar para novas localizacoes. O script valida os caminhos e executa `<exe> --version` antes de iniciar o pipeline; se algum falhar, regista em `log.txt` e aborta com mensagem clara.

## Pre-requisitos

1. Windows 11 com PowerShell 5.1+.
2. Microsoft Word e Outlook instalados (automacao COM).
3. Poppler e Tesseract instalados nos caminhos definidos acima (ou actualize o `config.ini`).
4. Permissoes para executar scripts (`Set-ExecutionPolicy RemoteSigned`).

## Como usar

1. Abra o PowerShell (a partir da pasta `PROCESSOS_GEP_Automacao`).
2. Execute: `./processar_processo.ps1`
3. Escolha o PDF do processo (`apn_XXXX.pdf`).
4. Aguarde cada passo. No final e apresentada uma mensagem de sucesso ou erro.

## O que o script faz

1. Cria estrutura do processo em `H:\PROCESSOS_GEP\<id>` e inicializa `log.txt`.
2. Copia o PDF para `origem`.
3. Executa pre-flight dos executaveis (`--version`) e verifica caminhos.
4. Converte o PDF para texto (Poppler) ou, em alternativa, usa Tesseract OCR.
5. Extrai campos (Nome, Morada, Codigo Postal, etc.) e grava `trabalho\data.json`.
6. Le o DOCX opcional (`peritagem_descricao_causas_conclusoes.docx`) e extrai secoes.
7. Gera `relatorio_fillform.pdf` via automacao do Word.
8. Cria um rascunho no Outlook com base no template HTML.
9. Regista tudo em `log.txt`.

## Logs e erros

- `trabalho\log.txt` inclui timestamps, pre-flight de versoes, mensagens de sucesso/erro e avisos.
- Se ocorrer algum erro critico (ex. falha na conversao), o script apresenta uma mensagem e e terminado.

## Personalizacao

- Ajuste o template de e-mail em `templates/email_template.html`.
- Os regex de extracao podem ser afinados na funcao `Extract-Fields` do script.
- O hook `fillform_placeholder.txt` permite adicionar futuramente a geracao de `fillform.json`.

## Teste manual do pdftotext

```
"H:\Programas instalados\Ferramentas\Poppler\bin\pdftotext.exe" -layout -nopgbrk -enc UTF-8 "H:\PROCESSOS_GEP\apn_XXXX\origem\apn_XXXX.pdf" "H:\PROCESSOS_GEP\apn_XXXX\trabalho\teste.txt"
```

Verifique `teste.txt` para confirmar que o texto esta legivel.

## Proximos passos possiveis

- Implementar geracao real de `fillform.json`.
- Ligar ao portal GEPProperty via Power Automate/Playwright.
- Botao "Fechar Processo" para arquivar e enviar resultados finais.
