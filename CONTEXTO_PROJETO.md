# Conversor de extrato bancario PDF -> Excel - Jean Vieira

## O que e o projeto

Ferramenta offline (sem IA, sem API, sem custo) que le extratos bancarios em
PDF de bancos e cooperativas usados pelos clientes de Jean, contador
(igrejas evangelicas e associacoes do terceiro setor) e gera, pra cada PDF:
um Excel com 4 colunas fixas (Data dd/mm/aaaa, Historico, Valor sempre
positivo com virgula decimal, Tipo C ou D) e um `.ofx` por banco pronto pra
importar no sistema de conciliacao bancaria.

Motivo: o cliente (Jean, contador) atualmente faz isso colando o PDF num
prompt de IA (Gemini) e copiando o resultado a mao pra planilha, depois
colando os valores no OFX Generator pra gerar o OFX. Objetivo e eliminar
essas etapas manuais com um programa que ele roda com dois cliques.

## Arquitetura

- `converter.py`: ponto de entrada. Le o(s) PDF(s) passado(s) como argumento,
  identifica o banco de cada pagina (funcao `identificar_banco`, lista
  `DETECTORES`), agrupa paginas consecutivas do mesmo banco
  (`extrair_blocos_por_banco`), chama o parser correspondente de `bancos.py`,
  e gera um `.xlsx` com **uma aba por banco** (nao mistura bancos na mesma
  aba, e nao coloca o nome do banco dentro do texto do Historico) mais uma
  aba "Pendencias" com avisos gerais.
- `bancos.py`: um `parse_<banco>(texto)` por banco, cada um retornando
  `(lista_de_transacoes, aviso_ou_None)`. Cada transacao e um dict com
  `data` (objeto `date`), `historico` (str), `valor` (Decimal, sempre
  positivo), `tipo` ('C' ou 'D'), `obs` (string, geralmente vazia).
- `ofx_export.py`: `gerar_ofx(transacoes, banco, caminho_saida)` gera um
  `.ofx` (formato OFX 1.02 SGML) a partir da mesma lista de transacoes usada
  no Excel. Formato validado batendo **byte a byte** contra um `.ofx` real
  exportado do OFX Generator (ferramenta que Jean usava manualmente antes
  disso) pro mesmo extrato Sicoob - por isso o cabecalho, a ausencia de
  bloco LEDGERBAL, o `CURDEF>USD` (nao BRL - e assim no arquivo de
  referencia) e o formato de `FITID`/`CHECKNUM` (`aammdd` + sequencial de
  2 digitos que reinicia a cada dia) nao devem ser alterados sem novo
  arquivo de referencia pra comparar. `BANKID` sai preenchido sozinho via
  `COMPE_BANCO` (codigo Febraban de 4 digitos, com zero a esquerda; 0756
  do Sicoob confirmado contra o arquivo de referencia, os demais sao a
  tabela publica Febraban). `ACCTID` sai fixo em `00000` (numero de conta
  nao da pra extrair do PDF de forma confiavel; usuario edita manualmente
  se o sistema de conciliacao exigir o numero real). Um `.ofx` por banco
  encontrado no PDF, nomeado `<nome_do_pdf>_<Banco>.ofx`.
- `Converter_Extratos.bat`: roda `python converter.py` em todos os PDFs da
  pasta onde o .bat esta, pra nao precisar usar linha de comando.
- `Para_Equipe/Conversor_Extratos.exe`: executavel standalone (PyInstaller,
  `--onefile`) gerado a partir de `converter.py`, pra distribuir pra equipe
  sem precisar instalar Python. Aceita duplo clique (processa todo PDF da
  pasta onde o .exe esta) ou arrastar PDFs em cima do icone (processa so os
  arrastados). Ver `Para_Equipe/LEIA-ME.txt` pras instrucoes que vao pra
  equipe.
- `Para_Equipe/LICENSE.txt`: licenca MIT (uso livre, "como esta", sem
  garantia), pra distribuicao mais ampla (nao so o time direto do Jean).

## Como gerar/atualizar o executavel

Sempre que `converter.py` ou `bancos.py` mudar e for hora de distribuir uma
nova versao pra equipe:

```
python -m pip install pyinstaller   # so na primeira vez
python -m PyInstaller --onefile --name Conversor_Extratos --console converter.py
```

Isso gera `dist/Conversor_Extratos.exe`. Copiar esse arquivo por cima do
`Para_Equipe/Conversor_Extratos.exe` pra atualizar a versao da equipe. O
`Conversor_Extratos.spec` guarda a configuracao do build (nao precisa mexer
nele pra um build padrao).

## Regra de ouro de cada parser

Nunca advinhar layout de banco sem ver uma amostra real. Sempre que um
parser for escrito ou ajustado, validar o resultado batendo a soma dos
creditos e a soma dos debitos contra os totais que o proprio extrato declara
(normalmente aparecem como "Total", "Saldo do periodo" ou similar). Isso ja
pegou varios bugs sutis de extracao de texto de PDF (linhas quebradas,
colunas que se misturam etc).

Armadilha mais comum: a extracao de texto de PDF (pdfplumber) as vezes
quebra uma linha visual em 2 ou 3 linhas de texto (descricao antes ou depois
do valor), ou junta coisas de forma inconsistente dentro do MESMO documento
(ja aconteceu do Bradesco ter um padrao numa pagina e outro padrao 2 paginas
depois). Sempre inspecionar o texto bruto (`page.extract_text()`) antes de
escrever o regex, e testar contra o extrato inteiro, nao so a primeira
pagina.

## Status por banco (atualizado nesta sessao)

### Validados com confianca total (bate exatamente com os totais do extrato)
- Sicoob
- Caixa Economica Federal (Gerenciador Caixa)
- Itau (excluir linhas de "Aplic Aut Mais", que sao varredura automatica
  pra CDB e nao contam como movimento de conta corrente conforme o proprio
  extrato)
- UniCred
- Cresol - dois formatos completamente diferentes, `parse_cresol` detecta
  qual e e roteia pro parser certo:
  1. Extrato tradicional (descricao as vezes quebra em 2 linhas com a linha
     de data+valor no meio).
  2. Extrato do sistema web "Colmeia" (`_parse_cresol_colmeia`), identificado
     por `EXTRATO CONSOLIDADO DE CONTA CORRENTE` ou pela URL
     `sistema.confesol/colmeia` no texto. Esse formato tem um problema
     serio de extracao: o `page.extract_text()` padrao do pdfplumber
     EMBARALHA a pagina inteira (parece ser peculiaridade de como esse
     sistema gera o PDF por "impressao" de pagina web). A solucao foi um
     fallback em `converter.py` (`_texto_robusto_por_caracteres`) que
     reconstroi as linhas agrupando os caracteres pela posicao vertical
     (`char['top']`) na ordem em que aparecem no fluxo do PDF, em vez de
     confiar no algoritmo de layout do pdfplumber - os caracteres em si
     estao em ordem de leitura correta no PDF, so o agrupamento em linhas
     que falha. O fallback so ativa quando detecta a URL
     `sistema.confesol/colmeia` no texto (ver `_FINGERPRINTS_TEXTO_EMBARALHADO`).
     Cada pagina tem tambem uma marca d'agua diagonal (nome de quem gerou o
     extrato, repetida varias vezes) que vira uma sopa de fragmentos curtos
     nesse agrupamento - `_parse_cresol_colmeia` filtra isso (linhas curtas
     demais, linhas com a URL do sistema, e usa o fato de que toda
     categoria real de transacao vem em CAIXA ALTA pra nao deixar um resto
     de marca d'agua sobrescrever a categoria certa quando uma transacao
     cai bem na quebra de pagina). Ver `EXTRATO CONSOLIDADO...pdf` usado
     pra validar - saldo inicial + soma dos lancamentos bateu exatamente
     com o saldo final declarado no extrato (13.474,42).
  IMPORTANTE: por causa desse formato novo, o detector do Bradesco em
  `DETECTORES` (`converter.py`) tinha um fallback frouxo (`'bradesco' in
  texto.lower()`) que dava falso positivo quando o extrato de OUTRO banco
  tinha um boleto pago pra "Bradesco Seguros" (nome de terceiro, nao do
  banco). Corrigido colocando `cresol` antes de `bradesco` na lista
  `DETECTORES` (ordem importa - `identificar_banco` para no primeiro
  detector que bate). Se aparecer um bug parecido com outro banco no
  futuro, suspeitar do mesmo padrao: fallback frouxo tipo `'nome_banco' in
  texto.lower()` pegando mencao a esse banco como texto livre dentro da
  descricao de uma transacao de outro banco.
- Ailos / ViaCredi (mesmo sistema, cooperativa aparece no cabecalho)
- Banco do Brasil - dois formatos, `parse_bb` detecta qual e:
  1. Formato atual (com colunas "Ag. origem"/"Lote" no cabecalho) - excluir
     linhas "BB Rende Facil", mesma logica do Itau.
  2. Formato de 2016 (`_parse_bb_2016`), sem "Ag. origem"/"Lote". Saldo so
     aparece na ULTIMA linha de um grupo de lancamentos do mesmo dia
     (todas as linhas anteriores do grupo ficam sem saldo), e "Documento"
     se distingue de "Valor" por nao ter virgula decimal. Descricao pode
     ter uma 2a linha de continuacao (ex: "SEFAZ RECURSOS ORDINARIOS")
     sem data na frente - cuidado pra nao confundir com a linha de saldo
     final "S A L D O" (tambem sem "documento" antes do valor, mas
     comeca com data, entao nao deve ser tratada como continuacao).
  ARMADILHA: os dois formatos tem "Dt. movimento" E "Dt. balancete" no
  cabecalho (so a ordem dos dois muda) - NAO da pra distinguir por isso
  sozinho (peguei uma regressao no formato atual testando so com isso).
  O sinal confiavel e a AUSENCIA de "Ag. origem" pra identificar o
  formato de 2016.
- Bradesco (cuidado: o layout dos primeiros lancamentos de credito difere
  do layout dos lancamentos de debito dentro do MESMO extrato - ver
  `parse_bradesco` pra entender a logica de linha "tipo" + linha de
  valor + continuacao)
- Sicredi

### Funciona mas precisa de revisao humana ocasional
- Santander: o texto extraido do PDF NAO tem coluna ou sinal confiavel de
  credito x debito (o layout usa duas colunas visuais que se perdem na
  extracao de texto simples). A classificacao e feita por palavra-chave da
  descricao (RECEBIDO/CREDITO = credito, PAGAMENTO/DEBITO/TARIFA = debito).
  Linhas incertas vem marcadas `[A VERIFICAR]` na coluna Observacao.

### Sem nenhuma amostra ainda (dos 14 bancos da carteira)
- Banco Civia
- Banrisul
- Sulcredi
- Votorantim

## Proximo passo

Quando o usuario mandar uma amostra completa (nao cortada) de algum desses
4 bancos, seguir o mesmo processo: extrair texto bruto com pdfplumber,
identificar o padrao de linha pra data/descricao/valor/tipo, escrever
`parse_<banco>` em `bancos.py`, adicionar o detector em `DETECTORES` no
`converter.py`, testar e validar contra os totais do extrato.

## Interface grafica (GUI) - em teste

Motivo: a tela estilo DOS do `.exe` console estava assustando usuarios
menos tecnicos (relatado pelo Jean, sessao 2026-09-08). Escolhida a opcao
`pywebview` (janela nativa reaproveitando o mesmo visual "ledger" usado em
`docs/index.html`) em vez de `tkinter`, por ficar bem mais bonita sem
precisar reescrever o design.

- `gui.py`: janela pywebview. Classe `Api` expoe `escolher_pdfs()`
  (dialogo nativo de arquivo) e `converter(caminhos)` (chama o mesmo
  `processar_pdf` do `converter.py` - a logica de conversao nao mudou,
  so a "casca") pro JS via `window.pywebview.api.*`.
- `gui.html`: interface (dropzone, lista de arquivos, botao Converter,
  lista de resultados com pill verde/vermelho). Mesma paleta de cores do
  `docs/index.html`, mas layout de ferramenta (nao de pagina de
  apresentacao) - sem tema escuro de proposito, e um app so, nao artifact.

**ARMADILHA IMPORTANTE - pywebview 6.x trava a janela**: a versao 6.2.1
(a mais recente no momento) tem um bug real (nao so log poluido) na
integracao com WebView2 nesta maquina - ao abrir a janela, tenta
enumerar `window.native.AccessibilityObject...` por reflexao e entra em
recursao infinita, deixando a janela com "Nao esta respondendo". A
correcao foi fixar a versao: `pip install pywebview==4.4.1` (sem esse
recurso experimental `window.native`, mais estavel). **Nunca fazer
`pip install --upgrade pywebview` sem testar antes que a janela abre e
continua respondendo (`Get-Process ... | Select Responding` no
PowerShell) por pelo menos uns 10 segundos** - o travamento nao aparece
imediato no log, só quando o processo trava mesmo.

Drag-and-drop de arquivo pra DENTRO da janela ja aberta (`file.pywebviewFullPath`
no JS) testado e confirmado que NAO funciona nesta maquina/versao do
WebView2 Runtime - cai no aviso "Nao consegui ler o caminho do arquivo
arrastado" (comportamento esperado, ver `dropzone.addEventListener('drop', ...)`
em `gui.html`). Por isso tem dois caminhos que SAO garantidos:
1. Botao "clique pra escolher" (dialogo nativo, `Api.escolher_pdfs`).
2. Arrastar o(s) PDF(s) em cima do ICONE do `.exe` no Explorer (nao pra
   dentro da janela) - o Windows chama o programa com os caminhos em
   `sys.argv`, igual o `.exe` de console ja fazia. `Api.arquivos_iniciais()`
   devolve esses caminhos pro JS assim que a janela carrega (evento
   `pywebviewready`), pré-preenchendo a lista. Os dois caminhos foram
   clicados e confirmados funcionando (nao so testado por logica).

**ARMADILHA (RESOLVIDA) - primeira execucao do .exe empacotado demorava
MUITO (10 a 30+ segundos) com a tela em branco**: onefile extrai tudo de
novo pra uma pasta temporaria TODA vez que abre (+ antivirus escaneando
os arquivos recem-extraidos de novo a cada vez), diferente de rodar
`python gui.py` direto (~7s, sem essa extracao). Corrigido com duas
mudancas, ambas testadas com controle de tela real (nao so por logica):
1. Build trocado de `--onefile` pra `--onedir` (pasta com o .exe e os
   DLLs ja extraidos, sem extracao repetida a cada abertura) - eliminou a
   demora na pratica, abriu rapido nas duas execucoes testadas.
2. `background_color='#EEF2E9'` no `webview.create_window()` (mesma cor
   de fundo do app) - pywebview usa branco puro por padrao, o que fazia
   qualquer atraso remanescente parecer tela travada. Sem isso o
   parametro default e `#FFFFFF`.

Build agora: `python -m PyInstaller --onedir --console --add-data
"gui.html;." gui.py` (gera uma PASTA em vez de um arquivo so - precisa
zipar junto com LEIA-ME.txt e LICENSE.txt igual o `.exe` de console,
nao da mais pra mandar um arquivo unico). Manter `--console` ate decidir
junto com o Jean se troca pra `--windowed`.

Status: testado e validado pelo Claude com controle de tela real (nao so
por logica) - janela renderiza, os dois fluxos de selecao de arquivo
funcionam, conversao bate com os totais corretos (testado com Sicoob 310,
Caixa 597, Itau 104 lancamentos, todos conferem). Falta o Jean decidir se
quer substituir o .exe de console pela GUI na distribuicao oficial (`Para_Equipe`,
release do GitHub) ou oferecer as duas opcoes.

## Preferencias do usuario (aplicam a este projeto)

- Portugues do Brasil, padrao numerico brasileiro (milhar com ponto,
  decimal com virgula).
- Nao usar travessao como pontuacao.
- Nao reconstruir de memoria um layout sem amostra real - pedir o PDF.
- Sempre confirmar contagem de lancamentos e totais batendo com o extrato
  antes de considerar um parser fechado.
