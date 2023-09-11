const xlsx = require('xlsx-populate'); // Importa a biblioteca xlsx-populate para trabalhar com arquivos Excel
const { createObjectCsvWriter } = require('csv-writer'); // Importa a biblioteca csv-writer para criar arquivos CSV
const iconv = require('iconv-lite'); // Importa a biblioteca iconv-lite para tratar caracteres especiais

function extrairDados(planilhaBuffer) {
  return new Promise((resolve, reject) => {
    xlsx
      .fromDataAsync(planilhaBuffer) // Carrega o arquivo Excel a partir do buffer de dados
      .then((workbook) => {
        const worksheet = workbook.sheet(0); // Obtém a primeira planilha do arquivo
        const data = worksheet.usedRange().value(); // Obtém os valores da planilha

        const registros = {
          docente: [], // Array para armazenar registros do tipo docente
          discente: [], // Array para armazenar registros do tipo discente
          outros: [], // Array para armazenar registros de outros tipos
        };

        for (let i = 0; i < data.length; i++) { 
          const email = data[i][0]; // Obtém o valor da coluna de email
          const primeiroNome = iconv.decode(data[i][1], 'ISO-8859-1'); // Obtém o valor da coluna de primeiro nome e faz a decodificação de caracteres especiais
          const ultimoNome = iconv.decode(data[i][2], 'ISO-8859-1'); // Obtém o valor da coluna de último nome e faz a decodificação de caracteres especiais
          const funcao = data[i][3]; // Obtém o valor da coluna de função
          let cpf = data[i][4]; // Obtém o valor da coluna de CPF
          // Tratamento do CPF
          if (cpf) {
            if (typeof cpf !== 'string') {
              // Verifica se é ou nao uma String
              cpf = String(cpf);
            }
            cpf = cpf.replace(/[^\d]/g, '').padStart(11, '0'); // Remove caracteres não numéricos do CPF e preenche com zeros à esquerda até completar 11 dígitos
            cpf = cpf.slice(0, 6); // Usa somente os 6 primeiros dígitos
          } else {
            cpf = '00000000000'; // Caso o CPF esteja vazio, atribui um valor padrão de 11 zeros
            cpf = cpf.slice(0, 6); // Usa somente os 6 primeiros dígitos
          }
          
          const registro = {
            Email: email,
            'Primeiro Nome': primeiroNome,
            'Último Nome': ultimoNome,
            CPF: cpf,
          };

        if (funcao.match(/^docente$/i)) {      // Aceitará o nome "docente" tanto em maiúsculo quanto minúsculo
          registros.docente.push(registro); // Adiciona o registro ao array de registros do tipo docente
        } else if (funcao.match(/^discente$/i)) { // Aceitará o nome "discente" tanto em maiúsculo quanto minúsculo
          registros.discente.push(registro); // Adiciona o registro ao array de registros do tipo discente
        } else {
          registros.outros.push(registro); // Adiciona o registro ao array de registros de outros tipos
        }
      }

        resolve(registros); // Retorna os registros extraídos da planilha
      })
      .catch((error) => reject(error)); // Rejeita a promessa em caso de erro
  });
}

function criarCSV(registros, tipo) {
  const nomeArquivo = `${tipo}.csv`; // Define o nome do arquivo com base no tipo de registro

  const csvWriter = createObjectCsvWriter({
    path: nomeArquivo, // Define o caminho do arquivo
    //header: registros[tipo].map(registro => registro.nome),
    header: [
      { id: 'Email'},
      { id: 'Primeiro Nome'},
      { id: 'Último Nome'},
      { id: 'CPF'},
    ], // Define o cabeçalho do arquivo CSV
  });

  return csvWriter
    .writeRecords(registros[tipo]) // Escreve os registros do tipo especificado no arquivo CSV
    .then(() => {
      console.log(`Arquivo ${nomeArquivo} criado com sucesso.`); // Exibe uma mensagem de sucesso
      return nomeArquivo; // Retorna o nome do arquivo criado
    })
    .catch((error) => {
      console.error(`Erro ao criar o arquivo ${nomeArquivo}:`, error); // Exibe uma mensagem de erro, caso ocorra algum problema na criação do arquivo
    });
}

module.exports = {
  extrairDados,
  criarCSV,
};
