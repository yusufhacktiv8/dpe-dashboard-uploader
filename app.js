const request = require('superagent');
const fs = require('fs');
const path = require('path');

const ExcelReader = require('./helpers/excel_reader');
const BadReader = require('./helpers/bad');
const UmurPiutangReader = require('./helpers/umur_piutang');
const CashFlowReader = require('./helpers/cash_flow');
const ProjectionReader = require('./helpers/projection');
const constant = require('./constant');

const signIn = (signInData) => {
  return new Promise((resolve, reject) => {
    postData(`${constant.serverUrl}/security/signin`, signInData)
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
      .set({'Content-Type': 'application/json', 'accept': 'json'})
      .end((err, res) => {
        if (!err) {
          resolve(res);
        } else {
          reject(err);
        }

      });
  });
};

const postDataSecure = (url, token, data) => {
  return new Promise((resolve, reject) => {
    console.log('Token: ', token);
    request
      .post(url)
      .send(data)
      .set({'Content-Type': 'application/json', 'Authorization': `Bearer ${token}`, 'accept': 'json'})
      .end((err, res) => {
        if (!err) {
          resolve(res);
        } else {
          reject(err);
        }

      });
  });
};

const displayResult = (result, title) => {
  return(`
    ==============================
    ${title}
    ==============================
    ${JSON.stringify(result)}
    ===============================
    `);
}

const processSend = (username, password, fileName, type, year, printCallback) => {
  signIn({ username, password })
  .then((token) => {
    if (!token) {
      printCallback('Login failed.');
      return;
    }

    if (type === 'OPS') {
      ExcelReader.readProjectProgress(fileName, year, (parseResult) => {
        postDataSecure(`${constant.serverUrl}/projectprogresses/batchcreate`, token, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'Project progress upload result'));
        })
        .catch((err) => {
          printCallback(displayResult(`Error: ${err.response.text}, status: ${err.response.status}`, 'Project progress upload error!'));
        });
      });

      ExcelReader.readLsp(fileName, year, (parseResult) => {
        postDataSecure(`${constant.serverUrl}/lsps/batchcreate`, token, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'LSP upload result'));
        });
      });

      ExcelReader.readClaim(fileName, year, (parseResult) => {
        postDataSecure(`${constant.serverUrl}/claims/batchcreate`, token, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'Claim upload result'));
        });
      });
    } else if (type === 'FIN1') {
      BadReader.readBad(fileName, year, (parseResult) => {
        postData(`${constant.serverUrl}/batchcreate/bad`, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'BAD upload result'));
        });
      });
    } else if (type === 'FIN2') {
      UmurPiutangReader.read(fileName, year, (parseResult) => {
        postData(`${constant.serverUrl}/batchcreate/umurpiutang`, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'Umur piutang upload result'));
        });
      });
    } else if (type === 'FIN3') {
      CashFlowReader.read(fileName, year, (parseResult) => {
        postData(`${constant.serverUrl}/batchcreate/cashflow`, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'Cashflow upload result'));
        });
      });
    }  else if (type === 'FIN4') {
      ProjectionReader.read(fileName, year, (parseResult) => {
        postData(`${constant.serverUrl}/batchcreate/projection`, parseResult)
        .then((res) => {
          printCallback(displayResult(res.body, 'Prognosa piutang upload result'));
        });
      });
    }
  });
}

exports.processSend = processSend;
