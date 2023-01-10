const UI = SpreadsheetApp.getUi();
const testingStrTo = 'jeffrey.storer@gmail.com';
const liveStrTo = makeLiveStrTo();

function test() {
  const INOUTWANT = getInOutWant('Want to Play', 26);
  console.log(INOUTWANT);
}

function getNamesOnWaitList() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Wait List');
  const range = sheet?.getDataRange();
  const values = range?.getValues();
  return values;
}

function getNameFromWaitList() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Wait List');
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const lastName = sheet.getRange(lastRow, 1).getValue();
    sheet.deleteRow(lastRow);
    return lastName;
  }
  return null;
}

function clearWaitList() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Wait List');
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.deleteRows(1, lastRow);
  }
}

function getLastNameFromEmail(email) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Addresses');
  const emailFinder = sheet.createTextFinder(email);
  const emailRange = emailFinder.findNext();
  if (!emailRange) return null;
  const row = emailRange.getRow();
  const lastName = sheet.getRange(row, 2).getValue();
  return lastName;
}

function getRowNumberFromValue(sheetName, rowValue, valueColumn) {
  //this is used with Sheet1 and Email Addresses
  let lastName = rowValue;
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(sheetName);
  const range = sheet.getDataRange();
  const textFinder = range.createTextFinder(lastName);
  const foundRange = textFinder.findNext();
  if (foundRange) {
    const rowNumber = foundRange.getCell(1, 1).getRow();
    return rowNumber;
  } else {
    return null;
  }
}

function handleEdit(e) {
  const rangeEdited = e.range;
  const sheetName = rangeEdited.getSheet().getName();
  const row = rangeEdited.getRow();
  switch (sheetName) {
    case 'In/Out':
    case 'Want to Play':
      processEdit(sheetName, row);
      break;
    case 'Sheet1':
      if (e.value === 'N') {
        checkForOpenings();
      }
      break;
    default:
      break;
  }
}

function checkForOpenings() {
  const openings = getNumberOfOpenings();
  if (openings > 0) {
    const playerName = getNameFromWaitList();
    if (playerName) {
      UI.alert('Name: ' + playerName);
      setYNCW(playerName, 'C');
      sendCWEmail(playerName, 'C');
    }
  }
}

function getLastName(sheetName, row) {
  //this is used with In/Out and Want to Play
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(sheetName);
  const email = sheet.getRange(row, 2).getValue();
  const lastName = getLastNameFromEmail(email);
  return lastName;
}

function getInOutWant(sheetName, row) {
  // this is used with In/Out and Want to Play
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(sheetName);
  const range = sheet.getDataRange();
  const value = range.getCell(row, 3).getValue();
  switch (value) {
    case 'In for Tomorrow':
      return 'In';
      break;
    case 'Out for Tomorrow':
      return 'Out';
      break;
    case 'Want to Play Tomorrow':
      return 'Want';
      break;
    default:
      return null;
      break;
  }
}

function processEdit(sheetName, row) {
  //get the last name
  const playerName = getLastName(sheetName, row);
  //check whether it's an in or out
  if (playerName) {
    const InOutWant = getInOutWant(sheetName, row);
    //if it's an out, set value to N
    //switch on sheetName
    switch (sheetName) {
      case 'In/Out':
        switch (InOutWant) {
          case 'In':
            //if it's an in, check whether value is Y
            const currentYNCW = getYNCW(playerName);
            if (currentYNCW === 'Y') {
              setYNCW(playerName, 'C');
            } else {
              //if not, treat it like a W request
              processWantEdit(playerName);
            }
            break;
          case 'Out':
            //get the last name
            const playerName = getLastName(sheetName, row);
            //set value to N
            setYNCW(playerName, 'N');
            break;
          default:
            break;
        }
        break;
      case 'Want to Play':
        switch (InOutWant) {
          case 'Want':
            //check if the value is a Y
            const currentYNCW = getYNCW(playerName);
            if (currentYNCW === 'Y') {
              //if so, set value to C
              setYNCW(playerName, 'C');
            } else {
              processWantEdit(playerName);
            }
            break;
          default:
            break;
        }
      default:
        break;
    }
  }
}

function processWantEdit(playerName) {
  //we get here if a player is a N but now wants to play
  const openings = getNumberOfOpenings();
  if (openings > 0) {
    setYNCW(playerName, 'C');
    sendCWEmail(playerName, 'C');
  } else {
    setYNCW(playerName, 'W');
    sendCWEmail(playerName, 'W');
    putNameOnWaitList(playerName);
  }
}

function getCWEmailStringTo(playerName) {
  const sheetName = 'Email Addresses';
  //get row number of playerName on email sheet
  const playerRow = getRowNumberFromValue(sheetName, playerName);
  //get value of email address
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = spreadSheet.getSheetByName(sheetName);
  const range = emailSheet.getDataRange();
  const emailAddress = range.getCell(playerRow, 1).getValue();
  return emailAddress;
}

function sendCWEmail(playerName, CW) {
  const strTo = getCWEmailStringTo(playerName);
  let strSubject;
  let strBody;
  switch (CW) {
    case 'C':
      strSubject = 'Wednesday Golf: You are in for tomorrow';
      strBody = 'We have room.  You are in.';
      break;
    case 'W':
      strSubject = 'Wednesday Golf: You are on the waiting list for tomorrow';
      strBody =
        'After everyone signed up has confirmed, I shall let you know whether we have an opening.  I should know by 6 p.m. today.';
      break;
    default:
      break;
  }
  GmailApp.sendEmail(strTo, strSubject, strBody);
}

function getScheduleSheetDataRange() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = spreadSheet.getSheetByName('Sheet1');
  const range = scheduleSheet.getDataRange();
  return range;
}

function getYNCW(playerName) {
  const playerRow = getRowNumberFromValue('Sheet1', playerName);
  const range = getScheduleSheetDataRange();
  const currentValue = range.getCell(playerRow, 4).getValue();
  return currentValue;
}

function setYNCW(playerName, value) {
  const playerRow = getRowNumberFromValue('Sheet1', playerName);
  const range = getScheduleSheetDataRange();
  range.getCell(playerRow, 4).setValue(value);
}

function getNumberOfOpenings() {
  const openingsRowNumber = getRowNumberFromValue('Sheet1', 'Openings');
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = spreadSheet.getSheetByName('Sheet1');
  const range = scheduleSheet.getDataRange();
  const openings = range.getCell(openingsRowNumber, 4).getValue();
  return openings;
}

function remStrBody() {
  let remPar1 =
    "If you want to play next Wednesday and have not already signed up, please click on the link below and enter 'Y' for at least next week's date before 6 p.m. today.\n\nIf you know your schedule beyond next week, you may enter that also.\n\nhttps://tinyurl.com/WednesdayGolfSchedules\n\nPlayers already signed up are: ";
  let remPar2 = getYWs('Y');
  let remPar3 =
    "\n\nPlease keep your tee choices for each course current: https://tinyurl.com/WednesdayGolfTeeChoices\n\nI shall put in a tee time request tonight for next week. If you don't sign up by 6 p.m., you will not be included in the request. You can be added to the booking later if there is room. Let me know.";
  let remPar4 = signature();
  let remPar = remPar1 + remPar2 + remPar3 + remPar4;
  return remPar;
}

function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function getLastRowInColumn(sheetName, row, column) {
  return getSheet(sheetName).getRange(row, column).getDataRegion().getLastRow();
}

function makeLiveStrTo() {
  const region = getSheet('Email Addresses').getRange(1, 1).getDataRegion();
  const playerCount = getLastRowInColumn('Email Addresses', 1, 1);
  let toString = '';
  let address = '';
  let i = 0;
  for (i = 2; i < playerCount; i++) {
    address = region.getCell(i, 1).getValue();
    toString = toString + address + ', ';
  }
  address = region.getCell(i, 1).getValue();
  return toString + address;
}

function getThisDayOfWeekNumber() {
  const d = new Date();
  return d.getDay();
}

function getDayOfWeekNumber(d) {
  return d.getDay();
}

function getMonthNumber(d) {
  return d.getMonth();
}

function getDayNumber(d) {
  return d.getDate();
}

function getTheHour(d) {
  return d.getHours();
}

function getTheMinutes(d) {
  return d.getMinutes();
}

function getDayOfWeekName(dayNumber) {
  let dayOfWeek = '';
  switch (dayNumber) {
    case 0:
      return (dayOfWeek = 'Sunday');
      break;
    case 1:
      return (dayOfWeek = 'Monday');
      break;
    case 2:
      return (dayOfWeek = 'Tuesday');
      break;
    case 3:
      return (dayOfWeek = 'Wednesday');
      break;
    case 4:
      return (dayOfWeek = 'Thursday');
      break;
    case 5:
      return (dayOfWeek = 'Friday');
      break;
    case 6:
      return (dayOfWeek = 'Saturday');
      break;
  }
}

function getMonthName(monthNumber) {
  let monthName = '';
  switch (monthNumber) {
    case 0:
      return (monthName = 'January');
      break;
    case 1:
      return (monthName = 'February');
      break;
    case 2:
      return (monthName = 'March');
      break;
    case 3:
      return (monthName = 'April');
      break;
    case 4:
      return (monthName = 'May');
      break;
    case 5:
      return (monthName = 'June');
      break;
    case 6:
      return (monthName = 'July');
      break;
    case 7:
      return (monthName = 'August');
      break;
    case 8:
      return (monthName = 'September');
      break;
    case 9:
      return (monthName = 'October');
      break;
    case 10:
      return (monthName = 'November');
      break;
    case 11:
      return (monthName = 'December');
      break;
  }
}
const linkRow =
  getSheet('Sheet1').getRange(1, 1).getDataRegion().getLastRow() + 4;
function playingDate() {
  const dateValue = getSheet('Sheet1')
    .getRange(linkRow + 4, 4)
    .getValue();
  const dayOfWeekNumber = getDayOfWeekNumber(dateValue);
  const dayOfWeekName = getDayOfWeekName(dayOfWeekNumber);
  const monthNumber = getMonthNumber(dateValue);
  const monthName = getMonthName(monthNumber);
  const dayNumber = getDayNumber(dateValue);
  return dayOfWeekName + ', ' + monthName + ' ' + dayNumber;
}

function course() {
  return getSheet('Sheet1')
    .getRange(linkRow + 5, 4)
    .getValue();
}

function linkTime() {
  const timeValue = getSheet('Sheet1')
    .getRange(linkRow + 6, 4)
    .getValue();
  const hour = getTheHour(timeValue);
  let minutes = getTheMinutes(timeValue).toString();
  if (minutes.length === 1) minutes = '0' + minutes;
  return hour + ':' + minutes;
}

function teeTimeCount() {
  return getSheet('Sheet1')
    .getRange(linkRow + 7, 4)
    .getValue();
}
function playerCount() {
  return getSheet('Sheet1')
    .getRange(linkRow + 8, 4)
    .getValue();
}
function waitCount() {
  return getSheet('Sheet1')
    .getRange(linkRow + 11, 4)
    .getValue();
}
function roomCount() {
  return getSheet('Sheet1')
    .getRange(linkRow + 12, 4)
    .getValue();
}

function makeStrSubject(dayOfWeek) {
  switch (dayOfWeek) {
    case 'Tuesday':
      return (
        'Wednesday Golf: Please Confirm for ' +
        playingDate() +
        ' at ' +
        course() +
        ' at ' +
        linkTime()
      );
      break;
    case 'Friday':
      return 'Wednesday Golf: Reminder to Sign Up for Next Week';
      break;
  }
}

function conIntro() {
  let conPar1 = 'We have ';
  if (teeTimeCount() > 1) {
    conPar1 = conPar1 + teeTimeCount() + ' tee times tomorrow at ';
  } else {
    conPar1 = conPar1 + teeTimeCount() + ' tee time tomorrow at ';
  }
  conPar1 =
    conPar1 +
    course() +
    ' starting at ' +
    linkTime() +
    '.  We have ' +
    playerCount() +
    ' players signed up';
  if (roomCount() > 0) {
    conPar1 = conPar1 + ' and room for ' + roomCount() + ' more.\n';
  } else {
    conPar1 = conPar1 + ' and no room for more.\n';
  }
  return conPar1;
}

function conWaitingCount() {
  switch (waitCount()) {
    case 1:
      return 'We have 1 player on the waiting list.\n';
      break;
    default:
      return 'We have ' + waitCount() + ' players on the waiting list.\n';
  }
}

function conSignedUp() {
  return 'Players signed up are: ' + getYWs('Y') + '\n';
}

function conWaitingList() {
  if (waitCount() > 0) {
    return 'Waiting List: ' + getYWs('W') + '\n';
  }
}

function conInOrOut() {
  return 'Who is in and who is out?\n';
}

function conLast() {
  return 'Please respond by 6 p.m..' + signature();
}

function signature() {
  return (
    '\n\n' +
    'Jeffrey B. Storer\n' +
    '617-279-6140 mobile\n' +
    'jeffrey.storer@gmail.com\n\n' +
    '53 Peregrine Crossing\n' +
    'Savannah, GA  31411-2863\n' +
    '912-335-1565'
  );
}

function getYWs(yw) {
  const region = getSheet('Sheet1').getRange(1, 1).getDataRegion();

  let count;

  switch (yw) {
    case 'Y':
      count = playerCount();
      break;
    case 'W':
      count = waitCount();
      break;
  }

  let getYWs = '';
  let j = 0;
  let i = 1;
  while (j < count) {
    i = i + 1;
    switch (count) {
      case 1:
        if (region.getCell(i, 4).getValue() === yw) {
          j = j + 1;
          getYWs = getYWs + region.getCell(i, 3).getValue();
        }
        break;
      case 2:
        if (region.getCell(i, 4).getValue() === yw) {
          j = j + 1;
          switch (j) {
            case 1:
              getYWs = getYWs + region.getCell(i, 3).getValue();
              break;
            case 2:
              getYWs = getYWs + ' and ' + region.getCell(i, 3).getValue();
              break;
            default:
          }
        }
        break;
      default:
        if (region.getCell(i, 4).getValue() === yw) {
          j = j + 1;
          switch (j) {
            case 1:
              getYWs = getYWs + region.getCell(i, 3).getValue();
              break;
            case count:
              getYWs = getYWs + ', and ' + region.getCell(i, 3).getValue();
              break;
            default:
              getYWs = getYWs + ', ' + region.getCell(i, 3).getValue();
          }
        }
    }
  }
  getYWs = getYWs + '.';
  return getYWs;
}

function makeStrBody(dayOfWeek) {
  switch (dayOfWeek) {
    case 'Tuesday':
      return makeConStrBody();
      break;
    case 'Friday':
      return makeRemStrBody();
      break;
  }
}

function makeConStrBody() {
  const br = '\n';
  let body = conIntro();
  if (waitCount() > 0) body = body + conWaitingCount();
  body = body + br + conSignedUp();
  if (waitCount() > 0) body = body + conWaitingList();
  return body + br + conInOrOut() + br + conLast();
}

function makeRemStrBody() {
  let body = remStrBody();
  return body;
}

function createMessage() {
  const htmlConIntro = conIntro();
  const htmlPar1 = htmlConIntro;
  const htmlConWaitingCount = conWaitingCount();
  const getHtmlPar2 = () => {
    if (waitCount() > 0) return htmlConWaitingCount;
    return '';
  };
  const htmlPar2 = getHtmlPar2();
  const htmlConSignedUp = conSignedUp();
  const htmlPar3 = htmlConSignedUp;
  const htmlConWaitingList = conWaitingList();
  const getHtmlPar4 = () => {
    if (waitCount() > 0) return htmlConWaitingList;
    return '';
  };
  const htmlPar4 = getHtmlPar4();
  let templ = HtmlService.createTemplateFromFile('ConfirmMessage.html');
  templ.htmlPar1 = htmlPar1;
  templ.htmlPar2 = htmlPar2;
  templ.htmlPar3 = htmlPar3;
  templ.htmlPar4 = htmlPar4;
  let message = templ.evaluate().getContent();
  return message;
}

function sendConfirmationOrReminder() {
  const props = {
    testing: false,
    day: '',
  };
  sendOut(props);
  return null;
}

function sendOut({ testing, day }) {
  let thisDay = getDayOfWeekName(getThisDayOfWeekNumber());
  let strTo = '';
  let strSubject = '';
  let strBody = '';

  if (testing) thisDay = day;

  switch (testing) {
    case true:
      strTo = testingStrTo;
      break;
    case false:
      strTo = liveStrTo;
  }

  if (thisDay !== 'Tuesday' && thisDay !== 'Friday') return null;
  strSubject = makeStrSubject(thisDay);
  strBody = makeStrBody(thisDay);
  if (thisDay === 'Friday') GmailApp.sendEmail(strTo, strSubject, strBody);
  if (thisDay === 'Tuesday') {
    const message = createMessage();
    GmailApp.sendEmail(strTo, strSubject, strBody, { htmlBody: message });
  }
}

function deleteCurrentWeek() {
  let sheet1 = getSheet('Sheet1');
  sheet1.deleteColumn(4);
  sheet1.getRange('B:B').activate();
  sheet1.getActiveRangeList().setBackground(null);
  addNewWeek();
}

function addNewWeek() {
  let sheetName = 'Sheet1';
  var sheet = getSheet(sheetName);
  let lastRow = getLastRowInColumn(sheetName, 1, 17);
  sheet.getRange(1, 17, lastRow, 1).activate();
  sheet.insertColumnsAfter(sheet.getActiveRange().getLastColumn(), 1);
  sheet
    .getActiveRange()
    .offset(
      0,
      sheet.getActiveRange().getNumColumns(),
      sheet.getActiveRange().getNumRows(),
      1
    )
    .activate();
  sheet.getRange(2, 18).activate();
  sheet.getCurrentCell().setValue('N');
  sheet
    .getActiveRange()
    .autoFill(
      sheet.getRange(2, 18, lastRow - 1, 1),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  sheet.getRange(1, 18).activate();
  let date = new Date(sheet.getRange(1, 17).getValue());
  date.setDate(date.getDate() + 7);
  sheet.getRange(1, 18).setValue(date);
  sheet.getRange(linkRow, 17, 13).activate();
  sheet
    .getActiveRange()
    .autoFill(
      sheet.getRange(linkRow, 17, 13, 2),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  sheet.getRange(linkRow, 17, 13, 2).activate();
  sheet.getRange(1, 4).activate();
}

function copyTeeChoicesToTLCGolfGoogleSheet() {
  const course = getSheet('Sheet1')
    .getRange(linkRow + 5, 4)
    .getValue();
  const sourceSheet = getSheet('Tee Choices');
  const courses = [
    'Deer Creek',
    'Magnolia',
    'Marshwood',
    'Oakridge',
    'Palmetto',
    'Terrapin Point',
  ];
  const courseIndex = courses.indexOf(course);
  if (courseIndex === -1) return;
  const courseColumn = courseIndex + 2;
  const playerCount =
    sourceSheet.getRange(1, 4).getDataRegion().getLastRow() - 1;
  const teeChoiceRange = sourceSheet.getRange(2, courseColumn, playerCount);
  const teeChoiceValues = teeChoiceRange.getValues();
  console.log(teeChoiceValues);
  const targetSpreadsheetId = '1GEP9S0xt1JBPLs3m0DoEOaQdwxwD8CEPFOXyxlxIKkg';
  const targetSheetName = '585871';
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  const dataRange = targetSheet.getRange(1, 1).getDataRegion();
  dataRange.activate();
  dataRange.createFilter();
  targetSheet.getRange('D1').activate();
  dataRange.getFilter().sort(4, false);
  const targetRange = targetSheet.getRange(2, 3, playerCount);
  targetRange.setValues(teeChoiceValues);
  targetSheet.getRange('B1').activate();
  dataRange.getFilter().sort(2, true);
  dataRange.getFilter().remove();
}

function testingRemind() {
  const props = {
    testing: true,
    day: 'Friday',
  };
  sendOut(props);
  return null;
}

function testingConfirm() {
  const props = {
    testing: true,
    day: 'Tuesday',
  };
  sendOut(props);
  return null;
}
