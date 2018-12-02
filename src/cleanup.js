var fs = require('fs');
var path = require('path');

function clearDir(dir) {
    var files = fs.readdirSync(dir);
    for (var i = 0; i < files.length; i++) {
        fs.unlinkSync(path.resolve(dir, files[i]));
    }
}

module.exports = function(outputDirCN, outputDirEN) {
    clearDir(outputDirCN);
    clearDir(outputDirEN);
}