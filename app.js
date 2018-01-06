const program = require('commander');
const co = require('co');
const prompt = require('co-prompt');
// const RSVP = require('rsvp');
const request = require('superagent');
const fs = require('fs');
const path = require('path');

program
  .option('-f, --filename <filename>', 'The user to authenticate as')
  .parse(process.argv);

  console.log('File name: ', program.filename);

  co(function *() {
    const username = yield prompt('Username: ');
    const password = yield prompt.password('Password: ');
    console.log('username', username);
    console.log('Password', password);
  });
