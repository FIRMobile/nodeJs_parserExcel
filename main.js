var alasql = require('alasql');
const XlsxPopulate = require('xlsx-populate');
var http = require("http");
    
const start = 'A1';
const end = 'G32';
const leti = {};
const timeLesson = {'8:00': [1], '9:50': [1], '11:40': [1], '13:45': [1], '15:35': [1], '17:25': [1], '19:05': [1]};
const dayWeek = {
    'ПОНЕДЕЛЬНИК': timeLesson,
    'ВТОРНИК': timeLesson,
    'СРЕДА': timeLesson,
    'ЧЕТВЕРГ': timeLesson,
    'ПЯТНИЦА': timeLesson,
    'СУББОТА': timeLesson,
    'ВОСКРЕСЕНЬЕ': timeLesson,
};
const excelFile = new Map();
let keyCells = {}; //ключевые места в разметке EXCEL-документа


XlsxPopulate.fromFileAsync("./test.xlsx")
.then(workbook => {
  for(let i=1; i<=31; i++)
    for(let j=66; j<=71; j++) {
      const coordinate = String.fromCharCode(j)+i;
      let value = workbook.sheet(0).cell(coordinate).value();
    
      if (value === '№ гр.' && !keyCells.numberGroupLine) {
        keyCells.numberGroupLine = {i, j};
        let column = j;
        while(true) {
          column++;
          let columnValue = workbook.sheet(0).cell(String.fromCharCode(column)+i).value();
          if (columnValue!==undefined) {
            leti[column] = {'name': columnValue, day: dayWeek};
          } else {
              break;
          }
        }
      }

      if (!!value) {
        if (String(value).indexOf('0.33') !== -1) {
            value = '8:00';
        }
        if (String(value).includes('0.40')) {
            value = '9:50';
        }
        if (String(value).includes('0.48')) {
            value = '11:40';
        }
        if (String(value).includes('0.57')) {
            value = '13:45';
        }
        if (String(value).includes('0.64')) {
            value = '15:35';
        }
        if (String(value).includes('0.72')) {
            value = '17:25';
        }
        if (String(value).includes('0.79')) {
            value = '19:05';
        }

        excelFile.set(i+' '+j, value);
      }
    }

    let mergeCells = workbook.sheet(0).cell("A1")._row._sheet._mergeCells;
    for (let key in mergeCells) {
      let startColum = key.split(':')[0];
      startColum = startColum[0];
      let startRow = Number(key.split(':')[0].replace(startColum, ''));

      let endColum = key.split(':')[1];
      endColum = endColum[0];
      let endRow = Number(key.split(':')[1].replace(endColum, ''));

      const coordinate = key.split(':')[0];
      let value = workbook.sheet(0).cell(coordinate).value();

      if (!!value) {
        if (String(value).indexOf('0.33') !== -1) {
            value = '8:00';
        }
        if (String(value).includes('0.40')) {
            value = '9:50';
        }
        if (String(value).includes('0.48')) {
            value = '11:40';
        }
        if (String(value).includes('0.57')) {
            value = '13:45';
        }
        if (String(value).includes('0.64')) {
            value = '15:35';
        }
        if (String(value).includes('0.72')) {
            value = '17:25';
        }
        if (String(value).includes('0.79')) {
            value = '19:05';
        }
      }

      for(let i=startRow; i<=endRow ; i++)
        for(let j=startColum.charCodeAt(0); j<=endColum.charCodeAt(0); j++) {
            const coordinate = String.fromCharCode(j)+i;
            excelFile.delete(i+' '+j);
            excelFile.set(i+' '+j, value);
        }
    }

    for(let i=keyCells.numberGroupLine.i+1; i<=31; i++)
      for(let j=keyCells.numberGroupLine.j+1; j<=71; j++) {
        if (excelFile.get(i+' '+j)) {
          const nameDay = excelFile.get(i+' '+String(keyCells.numberGroupLine.j-1));
          const keyLesson = excelFile.get(i+' '+keyCells.numberGroupLine.j);
          leti[j].day[nameDay][keyLesson].push(excelFile.get(i+' '+j));
        }
      }

    //console.log(excelFile['24 66']);
    console.log(leti);
});
/* alasql.promise('select * from XLSX("fkti-1-m.xlsx")')
        .then(function(data){
                console.log(data);
        }).catch(function(err){
                console.log('Error:', err);
        });
console.log("Сервер начал прослушивание запросов на порту 3000"); 
var XLSX = require('xlsx');
var workbook = XLSX.readFile('fkti-1-m.xlsx');
var sheet_name_list = workbook.SheetNames;
console.log('sheet_name_list', sheet_name_list);
sheet_name_list.forEach(function(y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var data = [];
    for(z in worksheet) {
        if(z[0] === '!') continue;
        //parse out the column, row, and value 
        console.log('Z', z);
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0,tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if(row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if(!data[row]) data[row]={};
        data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift();
    data.shift();
    console.log(data);
});*/