const { google } = require('googleapis');
const creds = require('./credentials.json');

const client = new google.auth.JWT(       // Autenticação para acesso à planilha.
  creds.client_email,
  null,
  creds.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

client.authorize(function(err, tokens){    // Verificação de algum erro p/ run da função assíncrona.
  if(err){
    console.log(err);
    return;
  } else {
      gsrun(client);
  }
});

let engSoftware = new Array(24);

async function gsrun(cl) {
  const gsapi = google.sheets({version: 'v4', auth: cl });    // Info do Google Sheet API para autenticação.

  var docId = '1rigrSjSthxusEtoDqGe7PEAjchl8nbuMTl1Fd1jV26w';   // Id do documento.

  const opt = {     // Coleta de dados das faltas.
    spreadsheetId: docId,
    range: 'engenharia_de_software!C4:C27'
  };
  let faltas = await gsapi.spreadsheets.values.get(opt);

  const opt1 = {    // Coleta de dados das notas da P1.
    spreadsheetId: docId,
    range: 'engenharia_de_software!D4:D27'
  };
  let p1 = await gsapi.spreadsheets.values.get(opt1);

  const opt2 = {    // Coleta de dados das notas da P2.
    spreadsheetId: docId,
    range: 'engenharia_de_software!E4:E27'
  };
  let p2 = await gsapi.spreadsheets.values.get(opt2);

  const opt3 = {    // Coleta de dados das notas da P3.
    spreadsheetId: docId,
    range: 'engenharia_de_software!F4:F27'
  };
  let p3 = await gsapi.spreadsheets.values.get(opt3);

  for(n=0; n<24; n++){    // Definição de um Array com todos os dados.
    engSoftware[n] =
      {
        faltas: parseInt(faltas.data.values[n][0]),
        p1: parseInt(p1.data.values[n][0]),
        p2: parseInt(p2.data.values[n][0]),
        p3: parseInt(p3.data.values[n][0]),
        media: 0,
        situacao: '',
        NAF: 0
      }
  }

  calcMedia();
  calcStatus();
  FillUS();

  console.log(engSoftware);
  console.log(UpdateSheet);

  const updateOptions = {   // Update na planilha.
    spreadsheetId: docId,
    range: 'engenharia_de_software!G4:H27',
    valueInputOption: 'USER_ENTERED',
    resource: { 
      values: UpdateSheet
    }
  };
  let res = await gsapi.spreadsheets.values.update(updateOptions);

}

function calcMedia() {    // Função para cálculo da média e da NAF.
  engSoftware.forEach( aluno => {
    aluno.media = Math.ceil((aluno.p1 + aluno.p2 + aluno.p3) / 3);
    if(aluno.media>=50 & aluno.media<70){
      aluno.NAF = 100 - aluno.media;
    } else {
        aluno.NAF = '0';
    }
  }) 
}

function calcStatus() {   // Função para cálculo da situação do aluno.
  for(n=0; n<24; n++)
    if(engSoftware[n].faltas>15) {
      engSoftware[n].situacao = 'Reprovado por falta';
      engSoftware[n].NAF = '0';
    } else if(engSoftware[n].media>=70) {
        engSoftware[n].situacao = 'Aprovado por nota';
    } else if(engSoftware[n].media<70 & engSoftware[n].media>=50) {
        engSoftware[n].situacao = 'Exame final';
    } else {
        engSoftware[n].situacao = 'Reprovado por nota'
    }
}

let UpdateSheet = new Array(24);

function FillUS(){      // Função para criar a matriz para Update.
  for(n=0; n<24; n++){
    UpdateSheet[n] = [engSoftware[n].situacao, engSoftware[n].NAF];
  }
}
