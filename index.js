const xlsx = require('xlsx');
const path = require('path');
const _ = require('lodash');
const moment = require('moment');
require('moment/locale/pt-br');

const file = path.resolve(__dirname, './input.xlsx');

let sheetData = null;
const data = {};
const startRow = 5;
const endRow = 100;
const ws_name = 'Registros';
let outputData = [];
let sheetDate = null;
let lastCol = 31;

// Available cols from 1 to 31
const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE']

/**
 * Get values at provided sheet, col and row
 * @param sheet
 * @param col
 * @param row
 * @returns {*}
 */
function at(sheet, col, row) {
  let v = sheet[`${col}${row}`] ? sheet[`${col}${row}`].v : null;
  v = typeof v === 'string' ? v.trim() : v;

  if (v === '*') return null;
  return v;
}

/**
 * Transform excel value to date
 * @param v
 * @param date1904
 * @returns {number}
 */
function datenum(v, date1904) {
  if(date1904) v+=1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

/**
 * Transform array of arrays into sheets for xlsx package
 * @param data
 * @param opts
 */
function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
    for(var C = 0; C != data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;
      var cell = {v: data[R][C] };
      if(cell.v == null) continue;
      var cell_ref = xlsx.utils.encode_cell({c:C,r:R});

      if(typeof cell.v === 'number') cell.t = 'n';
      else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
        cell.t = 'n'; cell.z = xlsx.SSF._table[14];
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if(range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range);
  return ws;
}


function Workbook() {
  if(!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function parseTime(times) {
  if (!times) return null;
  const timesList = `${times}`.split('\n');
  return timesList;
}

/**
 * Parse sheet data
 * @returns {null}
 */
function parseSheet() {
  let workbook = null;
  try {
    workbook = xlsx.readFile(file);
  } catch (e) {
    throw new Error(e);
  }

  sheetData = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheetData) return null;

  sheetDate = at(sheetData, 'C', 3).split(' ')[0];
  lastCol = at(sheetData, 'C', 3).split('~')[1].split('/')[1];

  let currentUser = null;
  for (let i = startRow; i <= endRow; i++) { //eslint-disable-line
    if (at(sheetData, 'A', i) === 'Nº') {
      currentUser = at(sheetData, 'K', i);
    } else {
      for (let day = 0, n = cols.length; day < n; day++) {
        const col = cols[day];
        const times = parseTime(at(sheetData, col, i));

        if (times) {
          data[currentUser] = data[currentUser] || {};
          data[currentUser][day] = times;
        }
      }
    }
  }
}

function getDurationBetweenDates(init, end) {
  const _entrance = moment.utc(init, 'HH:mm');
  const _exit = moment.utc(end, 'HH:mm');
  return moment.utc(+moment.duration(_exit.diff(_entrance))).format('HH:mm');
}

function sumDates(a, b) {
  const _a = moment(a, 'HH:mm');
  const _b = b.split(':');
  return _a.add(_b[0], 'hours').add(_b[1], 'minutes').format('HH:mm');
}

function minutesToTime(minutes) {
  const hours = minutes / 60;
  const h = _.round(hours - (hours % 1), 0);
  const m = _.round((hours % 1) * 60, 0);

  return `${h}:${m}`;
}

/**
 * Prepare parsed data to write
 */
function prepareSheetAndWrite() {
  const year = sheetDate.split('/')[0];
  const month = sheetDate.split('/')[1];
  const monthName = moment.months()[parseInt(month) - 1];

  _.map(data, (info, name) => {
    outputData = [['Dia', 'Dia da Semana', 'Entrada', 'Início Intervalo', 'Fim Intervalo', 'Saída', 'Carga Horária', 'Horas Trabalhadas']];
    let totalWorkedTime = 0;
    let totalWorkloadTime = 0;

    for (let i = 0; i < lastCol; i++) {
      const dayInfo = info[i] || null;

      let entrance = '';
      let breakInit = '';
      let breakEnd = '';
      let exit = '';
      let workload = '06:00';
      let workedHours = '';

      if (dayInfo) {
        if (dayInfo.length === 4) {
          entrance = dayInfo[0];
          breakInit = dayInfo[1];
          breakEnd = dayInfo[2];
          exit = dayInfo[3];

          const beforeBreak = getDurationBetweenDates(entrance, breakInit);
          const afterBreak = getDurationBetweenDates(breakEnd, exit);
          workedHours = sumDates(beforeBreak, afterBreak);

          const tmp = workedHours.split(':').map(i => parseInt(i));
          totalWorkedTime += (tmp[0] * 60) + tmp[1];
        } else {
          entrance = dayInfo[0];
          exit = dayInfo[1];

          if (entrance && exit) {
            workedHours = getDurationBetweenDates(entrance, exit);

            const tmp = workedHours.split(':').map(i => parseInt(i));
            totalWorkedTime += (tmp[0] * 60) + tmp[1];
          }
        }
      }

      const date = moment(new Date(parseInt(year), parseInt(month) - 1, i + 1, 0, 0, 0));
      if ([0,6].indexOf(date.weekday()) !== -1) {
        workload = '';
      } else {
        totalWorkloadTime += 6;
        workload = '06:00';
      }
      const row = [
        date.format('DD/MM/YYYY'),
        date.format('dddd'),
        entrance,
        breakInit,
        breakEnd,
        exit,
        workload,
        workedHours,
      ];
      outputData.push(row);
    }

    outputData.push([]);
    outputData.push(['', 'Total', '', '', '', '', minutesToTime(totalWorkloadTime * 60), minutesToTime(totalWorkedTime)]);

    const fileName = `${year}-${month} - ${monthName} - ${name}.xlsx`;
    writeSheet(outputData, fileName);
  })
}

/**
 * Write final sheet
 * @param output
 * @param fileName
 */
function writeSheet(output, fileName) {
  const wb = new Workbook();
  const ws = sheet_from_array_of_arrays(output);

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  /* write file */
  xlsx.writeFile(wb, fileName);
}

/**
 * Init
 */
function init() {
  parseSheet();
  prepareSheetAndWrite();
}

init();
