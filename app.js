const program = require('commander');
const co = require('co');
const prompt = require('co-prompt');
// const RSVP = require('rsvp');
const request = require('superagent');
const fs = require('fs');
const path = require('path');

const ExcelReader = require('./helpers/excel_reader');
const Constant = require('./Constant');

program
  .option('-f, --filename <filename>', 'The user to authenticate as')
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

// signIn({ username: 'yusuf', password: 'xupipuharo' })
signIn({ username: 'yusuf', password: 'admin' })
.then((token) => {
  ExcelReader.readProjectProgress(program.filename, (parseResult) => {
    postData(`${Constant.serverUrl}/batchcreate/projectprogress`, parseResult)
    .then((res) => {
      console.log(res);
    })
  });
});
