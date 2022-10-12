function randomize(month, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Firefighters'));
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var array = sheet.getRange('A2:A' + lastRow).getValues();
  let shuffled = shuffleAndCut(array);
  printNames(month, year, shuffled);
}

function confirmRandomize() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to randomize the truck checks?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var month = SpreadsheetApp.getActiveSheet().getRange('D9').getValue();
    var year = SpreadsheetApp.getActiveSheet().getRange('E9').getValue();;
    randomize(month, year);
  }
}

function shuffleAndCut(array) {
  let newArr = new Array(array.length);
  arrCopy = Array.from(array);
  var i = 0, j = 1;

  while(array.length > 0) {
    j = Math.floor(Math.random() * array.length);

    if (arrCopy[i] != array[j]) {
      let temp = "" + array[j];
      let last = temp.slice(temp.indexOf(" ") + 1);
      newArr[i] = last;
      array.splice(j, 1);
      i++;
    }
  }
  return newArr;
}

function printNames(month, year, shuffled) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName("" + month + " " + year));
  var currFiveWeeks = true;
  var printNamesFirstWeek = true;
  var arr = Array.from(shuffled);
  var backupArr = Array.from(shuffled);
  if (SpreadsheetApp.getActiveSheet().getRange('A6').isBlank()) {
    currFiveWeeks = false;
  }
  if (SpreadsheetApp.getActiveSheet().getRange('B2').getBackground() === '#000000') {
    printNamesFirstWeek = false;
  }
  if (printNamesFirstWeek && !currFiveWeeks) { //1 & 3
    var currOpenings = 18;
    var remaining = currOpenings - arr.length;
    while (remaining > 0) {
      arr.push(backupArr.shift());
      remaining--;
    }
    var i = 0, j = 2;
    while (arr.length > 0) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + j).setValue(arr.shift());
      SpreadsheetApp.getActiveSheet().getRange('' + col + j).setHorizontalAlignment('center');
      i++;
      if (col === 'J') {
        i = 0;
        j += 2;
      }
    }
  } else if (!printNamesFirstWeek) { // 2 & 4
    var currOpenings = 18;
    var remaining = currOpenings - arr.length;
    while (remaining > 0) {
      arr.push(backupArr.shift());
      remaining--;
    }
    var i = 0, j = 3;
    while (arr.length > 0) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + j).setValue(arr.shift());
      SpreadsheetApp.getActiveSheet().getRange('' + col + j).setHorizontalAlignment('center');
      i++;
      if (col === 'J') {
        i = 0;
        j += 2;
      }
    }
  } else if (printNamesFirstWeek && currFiveWeeks) { // 1 & 3 & 5
    var currOpenings = 27;
    var remaining = currOpenings - arr.length;
    while (remaining > 0) {
      arr.push(backupArr.shift());
      remaining--;
    }
  }
  var i = 0, j = 2;
  while (arr.length > 0) {
    var col = String.fromCharCode('B'.charCodeAt() + i);
    SpreadsheetApp.getActiveSheet().getRange('' + col + j).setValue(arr.shift());
    SpreadsheetApp.getActiveSheet().getRange('' + col + j).setHorizontalAlignment('center');
    i++;
    if (col === 'J') {
      i = 0;
      j += 2;
    }
  }
}

function newSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Firefighters'));
  var month = SpreadsheetApp.getActiveSheet().getRange('D2').getValue();
  var year = SpreadsheetApp.getActiveSheet().getRange('E2').getValue();
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  newSheet.setName("" + month + " " + year);
  format(month, year);
}

function format(month, year) {
  printTrucks();
  printDates(month, year);
  printBlack(month, year);
  randomize(month, year);
}

function printTrucks() {
  SpreadsheetApp.getActiveSheet().getRange('A1').setValue("Week Of/Truck");
  SpreadsheetApp.getActiveSheet().getRange('A1').setHorizontalAlignment('center');
  var i = 1, j = 1;
  while (j <= 9) {
    if (i == 3 || i == 5 || i == 9) {
      i++;
    }
    var col = String.fromCharCode('A'.charCodeAt() + j);
    SpreadsheetApp.getActiveSheet().getRange("" + col + '1').setValue(i);
    SpreadsheetApp.getActiveSheet().getRange("" + col + '1').setHorizontalAlignment('center');
    i++;
    j++;
  }
}

function printDates(month, year) {
  var monArr = getMondays(month, year);
  for (var i = 0; i < monArr.length; i++) {
    SpreadsheetApp.getActiveSheet().getRange('A' + (i + 2)).setValue("" + month + " " + monArr[i]);
  }
}

function getMondays(month, year) {
  var monNum = monthNumber(month);
  var days = new Date(parseInt(year, 10), parseInt(monNum,10) , 0).getDate();
  var mondays =  new Date(parseInt(monNum,10) +'/01/'+ parseInt(year, 10)).getDay();
  if (mondays != 1){
    mondays = 9 - mondays;
  }
  mondays = [mondays];
  for (var i = mondays[0] + 7; i <= days; i += 7) {
    mondays.push(i);
  }
  return mondays;
}

function printBlack(month, year) {
  var prev = previousMonth(month);
  var prevYear = 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prevSheet;
  var firstWeekPrintBlack = true;
  var prevMonthHadFive = true;
  var currMonthFive = false;
  if (prev === 'December') {
    prevYear = parseInt(year, 10) - 1;
    prevSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("" + prev + " " + prevYear));
    var last = ss.getActiveSheet().getLastRow();
  } else {
    prevSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("" + prev + " " + year));
    var last = ss.getActiveSheet().getLastRow();
  }
  SpreadsheetApp.setActiveSheet(ss.getSheetByName("" + month + " " + year));
  var cell = 'B' + last;
  if (prevSheet.getRange('B2').isBlank()) {
    firstWeekPrintBlack = false;
  }
  if (prevSheet.getRange('A6').isBlank()) {
    prevMonthHadFive = false;
  }
  if (!SpreadsheetApp.getActiveSheet().getRange('A6').isBlank()) {
    currMonthFive = true;
  }
  if (firstWeekPrintBlack && prevMonthHadFive) {
    var i = 0, row = 2, count = 0;
    while (count < 2) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + row).setBackground('black');
      i++;
      if (col === 'J') {
        i = 0;
        row += 2;
        count++;
      }
    }
  } else if (firstWeekPrintBlack && !prevMonthHadFive && currMonthFive) {
    var i = 0, row = 2, count = 0;
    while (count < 3) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + row).setBackground('black');
      i++;
      if (col === 'J') {
        i = 0;
        row += 2;
        count++;
      }
    } 
  } else if (firstWeekPrintBlack && !prevMonthHadFive && !currMonthFive) {
    var i = 0, row = 2, count = 0;
    while (count < 2) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + row).setBackground('black');
      i++;
      if (col === 'J') {
        i = 0;
        row += 2;
        count++;
      }
    }
  } else if (!firstWeekPrintBlack && !prevMonthHadFive && !currMonthFive) {
    var i = 0, row = 2, count = 0;
    while (count < 2) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + row).setBackground('black');
      i++;
      if (col === 'J') {
        i = 0;
        row += 2;
        count++;
      }
    }
  } else if (!firstWeekPrintBlack) {
    var i = 0, row = 3, count = 0;
    while (count < 2) {
      var col = String.fromCharCode('B'.charCodeAt() + i);
      SpreadsheetApp.getActiveSheet().getRange('' + col + row).setBackground('black');
      i++;
      if (col === 'J') {
        i = 0;
        row += 2;
        count++;
      }
    }
  }
}

function previousMonth(month) {
  if (month === 'January') {
    return 'December';
  }
  if (month === 'February') {
    return 'January';
  }
  if (month === 'March') {
    return 'February';
  }
  if (month === 'April') {
    return 'March';
  }
  if (month === 'May') {
    return 'April';
  }
  if (month === 'June') {
    return 'May';
  }
  if (month === 'July') {
    return 'June';
  }
  if (month === 'August') {
    return 'July';
  }
  if (month === 'September') {
    return 'August';
  }
  if (month === 'October') {
    return 'September';
  }
  if (month === 'November') {
    return 'October';
  }
  if (month === 'December') {
    return 'November';
  }
}

function monthNumber(month) {
  if (month === "January") {
    return 1;
  }
  if (month === "February") {
    return 2;
  }
  if (month === "March") {
    return 3;
  }
  if (month === "April") {
    return 4;
  }
  if (month === "May") {
    return 5;
  }
  if (month === "June") {
    return 6;
  }
  if (month === "July") {
    return 7;
  }
  if (month === "August") {
    return 8;
  }
  if (month === "September") {
    return 9;
  }
  if (month === "October") {
    return 10;
  }
  if (month === "November") {
    return 11;
  }
  if (month === "December") {
    return 12;
  }
}

