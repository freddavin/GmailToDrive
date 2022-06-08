// CÓDIGO PARA IMPLANTAÇÃO DO DASHBOARD

// Array de extensões no formato que se deseja baixar pro Drive
//var fileTypesToExtract = ['csv'];
// Nome da pasta no Drive onde serão baixado e manipulado os arquivos
var folderName = 'Dashboard Project';
// Nome da label no e-mail que será inserida após o download do arquivo
var labelName = 'saved';
// Constante para entensão do arquivo tratados na manipulação do Drive
var mimoType = 'text/csv'; // text/csv ou application/vnd.ms-excel (quando os arquivos são abertos pelo excel no windows, so funciona com essa extensão)
// id da pasta de download do drive, retirado na ultima parte da URL
var idFolderFiles = '1wcqsoXqvljBlSDlVspfCsFjyNSBPlOnt';

// metodo principal
function GmailToDrive() {
  // query que filtra os emails que serão inseridos na lista de emails a serem baixados o anexo.
  var query = 'in:inbox has:nouserlabels from:(lista-cdsdw@serpro.gov.br)';
  // o resultado na query é inserido numa lista de emails (lembrando que se um msm email tiver varias respostas, ou seja, uma conversa. Aqui ele entra como uma intancia só)
  var threads = GmailApp.search(query);
  // chama a função 'getGmailLabel_' para verificar se já possui uma label no email com o nome desejado e retoná-la. Se nao tiver, ela cria.
  var label = getGmailLabel_(labelName);
  // chama a função 'getFolder_' para verificar se já possui uma pasta no Drive com o nome desejado e retoná-la. Se nao tiver, ela cria.
  var parentFolder = getFolder_(folderName);

  if (threads.length > 0) {
    // função para fazer o backup dos arquivos que estão atualmente na pasta 'folderName', enviando pra uma subpasta definida pelo ID na função. E depois deleta os arquivos.
    backupAndDeleteFile();

    // recebe a pasta raiz do Drive
    var root = DriveApp.getRootFolder();    
    // loop pra baixar os arquivos do e-mail para a pasta definida.
    for (var i in threads) {
      // diferentemente do descrito na variavel threads, aqui ele verifica as conversas de um email e separa em uma lista de mensagens.
      var mesgs = threads[i].getMessages();
      for (var j in mesgs) {
        // para cada mensagem, a função 'getAttachments()' pega todos os anexos do email para fazer a conferencia de extensão
        var attachments = mesgs[j].getAttachments();
        for (var k in attachments) {
          // verifica anexo por anexo
          var attachment = attachments[k];
          // recebe uma copia do arquivo
          var attachmentBlob = attachment.copyBlob();
          // cria o arquivo na pasta raiz, pq o DriverApp não oferece um metodo para especificar o caminho
          var file = DriveApp.createFile(attachmentBlob);
          // coloca o arquivo na pasta desejada
          parentFolder.addFile(file);
          // remove da pasta raiz
          root.removeFile(file);
          Logger.log('Downloading file "%s".', file.getName());
        }
      }
      // coloca a label 'saved' no email em que realizou a copia do arquivo
      threads[i].addLabel(label);
    }

    // função para escrever o log de download com hora, data e nome do arquivo na aba 'Log' da planilha.
    writeLog();
    // função para converter o arquivo CSV do drive e escrever nas abas da planilha
    importCSVFile();

  } else {
    Logger.log('There is no files to download.');
  }
}

// função para verificar se já possui uma pasta no Drive com o nome desejado e retoná-la. Se nao tiver, ela cria.
function getFolder_(folderName) {
  var folder;
  // função retorna o endereço da pasta buscando pelo nome desejado.
  var fi = DriveApp.getFoldersByName(folderName);
  if (fi.hasNext()) {
    folder = fi.next();
  }
  else {
    // cria a pasta na raiz
    folder = DriveApp.createFolder(folderName);
  }
  return folder;
}

// função para verificar se já possui uma label no email com o nome desejado e retoná-la. Se nao tiver, ela cria.
function getGmailLabel_(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

// escreve o log de download de arquivos em uma aba da planilha chamada 'Log'
function writeLog() {
  // retorna a pasta desejada colocando o id dela (encontrado na URL do arquivo)
  var folderFiles = DriveApp.getFolderById(idFolderFiles);
  // retorna a pasta de trabalho (junção de todas as planilhas)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // retorna a planilha/aba desejada procurada pelo nome
  var logSheet = ss.getSheetByName("Log");

  // declara uma vetor para que será usada para incluir linha por linha, sendo cada coluna uma posição no vetor
  var results = [];
  // coloca na lista todos os arquivos da pasta buscando pelo tipo (que foi incluido como constante). É importante ressaltar novamente que neste momento, se for uma planilha de excel, o tipo de   arquivo tem que ser o outro especificado lá na constante 'mimoType', porém um não enxerga o outro se os arquivos estiverem misturados.
  var myfiles = folderFiles.getFilesByType(mimoType);

// limpa todo o conteudo da aba 'Log'.
  logSheet.clearContents();
  // escreve o cabeçalho novamente
  logSheet.appendRow(["Log Date", "File Date", "File Name"]);
  
  while (myfiles.hasNext()) {
    var file = myfiles.next();

    // variavel recebe o valor da hora atual, que será o momento em que fará o download e conversão
    var fdateLog = new Date(Date.now());
    // variavel recebe a data de geração do arquivo ou melhor dizendo a data de modificação dele
    var fdate = file.getLastUpdated();
    // variavel recebe o nome do arquivo
    var fname = file.getName();
    // coloca-se as informações das variaveis em um vetor
    results = [fdateLog, fdate, fname];
    // adiciona as informações na proxima linha vazia
    logSheet.appendRow(results);
  }
}

// função para organizar a planilha e converter os arquivos CSV
function importCSVFile() {
  // retorna a pasta desejada colocando o id dela (encontrado na URL do arquivo)
  var folderFiles = DriveApp.getFolderById(idFolderFiles);
  // retorna a pasta de trabalho (junção de todas as planilhas)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // coloca na lista todos os arquivos da pasta buscando pelo tipo (que foi incluido como constante).
  var files = folderFiles.getFilesByType(mimoType);

  // variavel de controle que será utilizada para verificação se deve ou não limpar a planilha que contem todos os 12 meses. 'DespesasGeraisDashboard'
  var firstTime = true;

  while (files.hasNext()) {
    var file = files.next();
    // variavel que recebera a quantidade de linhas que deverão ser deletadas. São as linhas do cabeçado original do csv do tesouro que serão substituidas
    var rows;
    // e suas colunas que serão deletadas
    var cols;
    // vetor que será informado o cabeçalho da planilha.
    var descCols;
    // nome da aba da planilha em que serão transferido as informações do CSV
    var nameSheet;
    // delimitador utilizado pelo CSV, no nosso caso foi definido a ',' no momento de configuração no tesouro direto.
    var delimitador;

    Logger.log('Converting CSV file "%s".', file.getName());

    // quando pegamos o nome do arquivo, ele vem de uma forma que precisamos cortar de maneira que consiga tratar o caso do 'LoaDashboard' maneira especifica
    if (file.getName().substring(0, 23) == 'LoaDashboard.csv') {
      // intanciando as variaveis declaradas anterioremnte / informações particulares de cada tipo de planilha
      rows = 1;
      cols = 14;
      descCols = [["Ação Governo Descrição", "Ação Governo", "Fonte de Recursos", "Fonte de Recursos Descrição", "Grupo Despesa Descrição",
        "Dotação Inicial", "Empenhado", "Crédito Disponível + Crédito Aguardando Limite Orç.", "Crédito Disponível", "Liquidado", "Pago", "A Liquidar inscritos em RPNP",
        "Liquidados inscritos em RPP", "Total"]];
      nameSheet = 'LoaReport';
      delimitador = ",";

      // retorna a planilha/aba desejada procurada pelo nome
      var sheet = ss.getSheetByName(nameSheet);
      // matriz usada para extrair linha por linha do CSV e colocar na planilha
      var csvArray;

      // limpa toda a planilha. Decidi incluir essa linha para ter certeza de que não fique residuos das informações anteriores, por mais que depois seja deletada.
      sheet.clear();
      // exclusão das linhas
      sheet.deleteRows(2, sheet.getMaxRows() - 1);
      // insere as informações anteriores no alcance desejado, que é da A1 aé A14
      sheet.getRange(1, 1, 1, cols).setValues(descCols);

      // insere na variavel a copia do arquivo convertida em string na codificação 'UTF-16'. Importante campo tb! Caso o arquivo trabalhado fique desconfigurado, verificar sua codificação.
      csvArray = file.getBlob().getDataAsString('UTF-16');
      // função da biblioteca Utilities, que insere a string desejada e o delimitador para conversão
      csvArray = Utilities.parseCsv(csvArray, delimitador);
      // neste momento o 'csvArray' é uma matriz em que as colunas 2, 3 e assim por diante estão vazias. Ou seja, uma matriz com dimensão N x 1 x Z, onde Z neste caso possui um vetor de 14 colunas.
      // Com isso precisamos pegar somente a primeira coluna e remover todo o resto. Que é oq a função shift faz.
      csvArray.shift();

      // recebe a posição da ultima linha
      var lastRow = sheet.getLastRow();
      // insere o numero de linhas necessarios para o tamanho do vetor, ou seja, do CSV
      sheet.insertRowsAfter(lastRow, csvArray.length);
      // insere as informações do vetor no alcance desejado, que é do A2 até o tamanho do vetor
      sheet.getRange(lastRow + 1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);

      // uma vez que a planilha gerada pelo tesouro tem os cabeçalhos desorganizados, é necessário deletar, pois ja criamos um cabeçalho personalizado. Então ele pega a linha 2 até o numero de linhas instanciados no inicio da condicional e deleta.
      sheet.deleteRows(lastRow + 1, rows);

    } else if (file.getName().substring(0, 23) == 'DespesasGeraisDashboard') {
      rows = 2;
      cols = 19;
      descCols = [["Fonte de Recursos", "Fonte de Recursos Descrição", "Ação Governo", "Ação Governo Descrição", "PTRES", "UGR", "UGR Descrição",
        "Grupo Despesa", "Grupo Despesa Descrição", "Mês Lançamento", "Natureza Despesa", "Natureza Despesa Descrição", "Favorecido", "Favorecido Descrição",
        "PI", "PI Descrição", "Item Informação", "Item Informação Descrição", "Movimento Líquido"]];
      nameSheet = 'DespGeraisTotal';
      delimitador = "\t";

      var sheet = ss.getSheetByName(nameSheet);
      var csvArray;

      // condicional que será aceita somente na primeira vez, fazendo os procedimentos para limpar e a planilha. Após isso ela vai jogar todos os arquivos no final de cada linha preenchida. Formando uma planilha só, com todas as informações.
      if (firstTime == true) {
        sheet.clear();
        sheet.deleteRows(2, sheet.getMaxRows() - 1);
        sheet.getRange(1, 1, 1, cols).setValues(descCols);
        firstTime = false;
      }

      csvArray = file.getBlob().getDataAsString('UTF-16');
      csvArray = Utilities.parseCsv(csvArray, delimitador);
      csvArray.shift();

      var lastRow = sheet.getLastRow();

      sheet.insertRowsAfter(lastRow, csvArray.length);
      sheet.getRange(lastRow + 1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);

      sheet.deleteRows(lastRow + 1, rows);

    } else {
      Logger.log('There is no file with the correct name to convert.');
    }
  }
}

// faz o backup dos arquivos antigos, renomeando o arquivo com a data de backup, e apaga da pasta principal
function backupAndDeleteFile() {
  var folderFiles = DriveApp.getFolderById(idFolderFiles);
  var backupFolder = DriveApp.getFolderById('1iijseCyjxvbgvRgB9AHZzNVF0jMbGJ-8');

  var files = folderFiles.getFilesByType(mimoType);

  while (files.hasNext()) {
    var file = files.next();
    var fdateLog = Utilities.formatDate(new Date(Date.now()), "GMT", "dd-MM-yyyy");
    Logger.log('Backup file "%s".', file.getName());
    file.makeCopy(fdateLog + ' ' + file.getName(), backupFolder);
    Logger.log('Deleting file "%s".', file.getName());
    Drive.Files.remove(file.getId());
  }
}













