const app = require('./app.js');

const initApp = function() {

  const theUsername = document.getElementById('theUsername');
  const thePassword = document.getElementById('thePassword');
  const theFile = document.getElementById('theFile');
  const theType = document.getElementById('theType');
  const theYear = document.getElementById('theYear');
  const theButton = document.getElementById('theButton');
  const theResult = document.getElementById('theResult');

  theButton.onclick = function(e) {
    const username = theUsername.value;
    const password = thePassword.value;
    const fileName = theFile.files.item(0).name;
    const filePath = theFile.files.item(0).path;
    const fileType = theType.value;
    const year = theYear.value;
    theResult.value = "Uploading file..."

    app.processSend(username, password, filePath, fileType, year, (processResult) => {
      theResult.value = theResult.value + '\n\n' + processResult;
    });
  };
}

exports.initApp = initApp;
