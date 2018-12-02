var fs = require('fs');
var path = require('path');
var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
function renderDoc(templatePath, outputPath, data) {

    var content = fs
        .readFileSync(templatePath, 'binary');

    var zip = new JSZip(content);

    var doc = new Docxtemplater();
    doc.loadZip(zip);

    doc.setData(data);
    doc.render()

    var buf = doc.getZip()
        .generate({ type: 'nodebuffer' });

    fs.writeFileSync(outputPath, buf, {
        flag: 'w'
    });
}

module.exports = function (cnTemplatePath, enTemplatePath, outputDirCN, outputDirEN, data) {
    var clientItem;
    var outputPath;
    var templatePath;
    for (var i = 0; i < data.length; i++) {
        clientItem = data[i];
        templatePath = clientItem.isDomestic ? cnTemplatePath : enTemplatePath;
        outputDir = clientItem.isDomestic ? outputDirCN : outputDirEN;
        outputPath = clientItem.clientName.replace(/[\/\\]/g, ' ');
        outputPath = path.resolve(outputDir, outputPath + '.docx');
        renderDoc(templatePath, outputPath, clientItem);
    }
}