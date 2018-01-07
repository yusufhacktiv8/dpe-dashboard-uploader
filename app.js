const program = require('commander');
const co = require('co');
const prompt = require('co-prompt');
// const RSVP = require('rsvp');
const request = require('superagent');
const fs = require('fs');
const path = require('path');

const ExcelReader = require('./helpers/excel_reader');
const BadReader = require('./helpers/bad');
const UmurPiutangReader = require('./helpers/umur_piutang');
const CashFlowReader = require('./helpers/cash_flow');
const ProjectionReader = require('./helpers/projection');
const Constant = require('./Constant');

program
  .option('-f, --filename <filename>', 'Filename')
  .option('-t, --type <type>', 'Type')
  .parse(process.argv);

  console.log('File name: ', program.filename);

  // co(function *() {
  //   const username = yield prompt('Username: ');
  //   const password = yield prompt.password('Password: ');
  //   console.log('username', username);
  //   console.log('Password', password);
  //
  //   request
  //     .post('http://dashboard-dpe.wika.co.id/api/security/signin')
  //     .send({ username, password }) // sends a JSON post body
  //     .set('accept', 'json')
  //     .end((err, res) => {
  //       console.log(res.body.token);
  //       ExcelReader.readProjectProgress(program.filename, (projectProgresses) => {
  //         console.log(projectProgresses);
  //       });
  //     });
  //   });

const signIn = (signInData) => {
  return new Promise((resolve, reject) => {
    postData(`${Constant.serverUrl}/security/signin`, signInData)
    .then((res) => {
      resolve(res.body.token);
    });
  });
};

const postData = (url, data) => {
  return new Promise((resolve, reject) => {
    request
      .post(url)
      .send(data)
      .set('accept', 'json')
      .end((err, res) => {
        resolve(res);
      });
  });
};

const displayResult = (result, title) => {
  console.log(`
    ==============================
    ${title}
    ==============================
    ${JSON.stringify(result)}
    ===============================
    `);
}

const type = program.type;

signIn({ username: 'yusuf', password: 'xupipuharo' })
// signIn({ username: 'yusuf', password: 'admin' })
.then((token) => {

  if (type === 'OPS') {
    ExcelReader.readProjectProgress(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/projectprogress`, parseResult)
      .then((res) => {
        displayResult(res.body, 'Project progress upload result');
      });
    });

    ExcelReader.readLsp(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/lsp`, parseResult)
      .then((res) => {
        displayResult(res.body, 'LSP upload result');
      });
    });
  } else if (type === 'FIN1') {
    BadReader.readBad(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/bad`, parseResult)
      .then((res) => {
        displayResult(res.body, 'BAD upload result');
      });
    });
  } else if (type === 'FIN2') {
    UmurPiutangReader.read(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/umurpiutang`, parseResult)
      .then((res) => {
        displayResult(res.body, 'Umur piutang upload result');
      });
    });
  } else if (type === 'FIN3') {
    CashFlowReader.read(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/cashflow`, parseResult)
      .then((res) => {
        displayResult(res.body, 'Cashflow upload result');
      });
    });
  }  else if (type === 'FIN4') {
    ProjectionReader.read(program.filename, (parseResult) => {
      postData(`${Constant.serverUrl}/batchcreate/projection`, parseResult)
      .then((res) => {
        displayResult(res.body, 'Prognosa piutang upload result');
      });
    });
  }

});
