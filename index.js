var path = require('path');
var cleanup = require('./src/cleanup.js');
var parse = require('./src/parse.js');
var render = require('./src/render.js');

var outputDirCN = path.resolve(__dirname, 'output/cn/');
var outputDirEN = path.resolve(__dirname, 'output/en/');
cleanup(outputDirCN, outputDirEN);
console.log('cleaned up!')
parse(path.resolve(__dirname, 'input/debit.xlsx'), 'debit', function(err, data) {
    if (err) {
        console.log(err);
        return;
    }

    // console.log('data: ', data);
    render(path.resolve(__dirname, 'templates/cn.docx'), path.resolve(__dirname, 'templates/en.docx'), outputDirCN, outputDirEN, data);
    console.log('okk!');
})