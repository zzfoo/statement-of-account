var exceljs = require('exceljs');
var config = require('./config.js');

function parseExcel(filePath, sheetName, callback) {
    var workbook = new exceljs.Workbook();
    var error;
    workbook.xlsx.readFile(filePath)
        .then(function() {
            // console.log('读取 ' + filePath + ' 成功！');
            var worksheet =  workbook.getWorksheet(1);
            if (!worksheet) {
                error = '找不到 excel 页 ' + sheetName;
                callback(error);
                return;
            }
            var headerKeys = config.headerKeys;
            var keys = [];
            worksheet.getRow(1).eachCell(function(cell) {
                var header = cell.value.trim();
                var key = headerKeys[header] || null;
                keys.push(key)
            });

            var items = [];
            worksheet.eachRow(function(row, rowNumber) {
                if (rowNumber === 1) {
                    return;
                }

                var item = {};
                row.eachCell(function(cell, cellNumber) {
                    var key = keys[cellNumber - 1];
                    if (key) {
                        item[key] = cell.value;
                    }
                })
                items.push(item);
            })

            callback(null, items);
        })
        .catch(function(err) {
            console.log('parseExcel: ', err);
            // error = '读取 ' + filePath + ' 错误！';
            callback(error);
        })
}

function groupByClient(data) {
    var foriegnCurrency = '美元';
    var foriegnCurrencySign = '$';
    var nativeCurrencySign = '￥';
    var clientItems = [];
    var placed = {};
    var clientName;
    var item;
    var clientItem;
    var isDomestic;
    var currencySign;
    for (var i = 0; i < data.length; i++) {
        item = data[i];
        clientName = item.clientName;
        if (placed[i]) {
            continue;
        }
        isDomestic = item['currency'] !== foriegnCurrency;
        currencySign = isDomestic ? nativeCurrencySign : foriegnCurrencySign;
        clientItem = {
            'clientName': clientName,
            'isDomestic': isDomestic,
            'invoices': [],
        };
        for (var j = i; j < data.length; j++) {
            if (data[j].clientName === clientName) {
                clientItem.invoices.push(data[j]);
                placed[j] = true;
            }
        }
        clientItem['debitTotal'] = 0;
        clientItem['debit0-30'] = 0;
        clientItem['debit31-60'] = 0;
        clientItem['debit61-90'] = 0;
        clientItem['debit91AndMore'] = 0;
        var time;
        var date;
        var now = new Date().setHours(0, 0, 0, 0);
        var day30 = 30 * 24 * 3600 * 1000;
        var date30 = now - day30;
        var date60 = date30 - day30;
        var date90 = date60 - day30;
        var fee;
        var year, month, day;
        for (var k = 0; k < clientItem.invoices.length; k++) {
            date = clientItem.invoices[k].createdDate;
            time = date.getTime();
            fee = clientItem.invoices[k].fee;
            clientItem['debitTotal'] += fee;
            if (time > date30) {
                clientItem['debit0-30'] += fee;
            } else if (time > date60) {
                clientItem['debit31-60'] += fee;
            } else if (time > date90) {
                clientItem['debit61-90'] += fee;
            } else {
                clientItem['debit91AndMore'] += fee;
            }
            clientItem.invoices[k].fee = currencySign + toFixed(fee);
            year = date.getYear() + 1900;
            month = date.getMonth() + 1;
            if (month < 10) {
                month = '0' + month;
            }
            day = date.getDate();
            if (day < 10) {
                day = '0' + day;
            }
            if (isDomestic) {
                clientItem.invoices[k].createdDate = year + '/' + month + '/' + day;
            } else {
                clientItem.invoices[k].createdDate = month + '/' + day + '/' + year;
            }
            if (!clientItem.invoices[k]['clientFileId'] || clientItem.invoices[k]['clientFileId'] === '无') {
                clientItem.invoices[k]['clientFileId'] = '';
            }
            if (!clientItem.invoices[k]['agentFileId'] || clientItem.invoices[k]['agentFileId'] === '无') {
                clientItem.invoices[k]['agentFileId'] = '';
            }
        }
        clientItem['debitTotal'] = currencySign + toFixed(clientItem['debitTotal']);
        clientItem['debit0-30'] = currencySign + toFixed(clientItem['debit0-30']);
        clientItem['debit31-60'] = currencySign + toFixed(clientItem['debit31-60']);
        clientItem['debit61-90'] = currencySign + toFixed(clientItem['debit61-90']);
        clientItem['debit91AndMore'] = currencySign + toFixed(clientItem['debit91AndMore']);
        clientItems.push(clientItem);
    }
    return clientItems;
}

function toFixed(number) {
    return ((number * 100) >> 0) / 100;
}

module.exports = function(filePath, sheetName, callback) {
    parseExcel(filePath, sheetName, function(err, data) {
        if (err) {
            callback(err);
            return;
        }

        data = groupByClient(data);
        callback(null, data);
    })
}