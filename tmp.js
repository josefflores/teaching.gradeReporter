/**
 * Jose Flores
 * jose.flores.152@gmail.com
 * 10/20/16
 */

// Defaults
var g_debug_val = 'false';
var appName = 'Grade reporter';

// Global Variables
var g_debug_key = 'g_debug';
var g_debugEmail_key = 'g_debugEmail';
var userProp = PropertiesService.getUserProperties(); // user settings

/**
 * Adds the Teaching menu to to the spreadsheet UI at initialization.
 */
function onOpen() {
  setDefaults();
  loadMenu(userProp.getProperty(g_debug_key));
}

/**
 * Loads the Teaching menu. It has two states debug and production.
 *
 * @param debug string   The string representation of true or false,
 *                       'true' display debug menu
 *                       'false' display normal menu
 */
function loadMenu(debug){
  // Generate program options
  var menu = SpreadsheetApp.getUi()
        .createMenu('Teaching')
        .addItem('Send grades to all student rows', 'sendGradesAll')
        .addItem('Send grade to individual student by row', 'sendGradesSelect')
        .addSeparator();

  // Generate Debug options
  if(debug == 'true') {
      menu.addItem('Turn debug off', 'turnOffDebug')
        .addItem('Change debug email <' + userProp.getProperty(g_debugEmail_key) + '>', 'changeDebugEmail')
        .addItem('Reset debug defaults', 'resetDebug');
  } else {
      menu.addItem('Turn debug on', 'turnOnDebug');
  }

  // Update UI
  menu.addToUi();
}

/**
 * Sets sheet properties to default values.
 */
function setDefaults(){
  // Default debug email from current user
  if(userProp.getProperty(g_debugEmail_key) === null)
      userProp.setProperty(g_debugEmail_key, Session.getActiveUser().getEmail());

  // Default debug state
  if(userProp.getProperty(g_debug_key) === null)
      userProp.setProperty(g_debug_key, g_debug_val);
}

/**
 * Menu Action - Reset sheet properties, and reloads menu with changes.
 */
function resetDebug() {
  userProp.deleteProperty(g_debugEmail_key);
  userProp.deleteProperty(g_debug_key);
  setDefaults();
  loadMenu(userProp.getProperty(g_debug_key));
}

/**
 * Menu Action - Turns on debug mode, and reloads menu with changes.
 */
function turnOnDebug(){
  userProp.setProperty(g_debug_key, 'true');
  loadMenu(userProp.getProperty(g_debug_key));
}
/**
 * Menu Action - Turns off debug mode, and reloads menu with changes.
 */
function turnOffDebug(){
  userProp.setProperty(g_debug_key, 'false');
  loadMenu(userProp.getProperty(g_debug_key));
}

/**
 * Menu Action - Kicks off the debug email change, prompts user for input, validates, and reloads menu with changes.
 */
function changeDebugEmail(){
  /**
   * Validates an email for structure.
   *
   * @param email string     The email to validate
   * @return      bool       true - valid email
   *                         false - invalid email
   */
  var validateEmail = function(email) {
    // Valid email regex
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
  }

  var query, email,
      sheet = SpreadsheetApp.getActiveSheet(),
      ui = SpreadsheetApp.getUi();

      // Prompt for email.
      query = ui.prompt(appName,
                        'Enter new debug email.',
                        ui.ButtonSet.OK_CANCEL);

      // Email is stored.
      email = query.getResponseText();

  // Handle prompt result
  if (query.getSelectedButton() == ui.Button.OK) {
    // Email is valid
    if (validateEmail(email)) {
      userProp.setProperty(g_debugEmail_key, email);
      loadMenu(userProp.getProperty(g_debug_key));
      return success(appName, 'Debug email was changed.', userProp.getProperty(g_debugEmail_key));
    }
    // Email is invalid
    return fail(appName, 'Invalid email.');
  }
  // Canceled Action
  return fail(appName, 'Debug email change was cancelled.');
}

/**
 * Shorthand for getting value of cell.
 *
 * @param    sheet    <Sheet>   The sheet to extract from.
 * @param    range    <string>  The range to extract.
 * @return   The value of the cell.
 */
function cell(sheet, range){
  return sheet.getRange(range)
      .getCell(1,1)
      .getValue();
}

/**
 * Wrapper function to send all grade reports.
 */
function sendGradesAll(){
  sendGrades(true);
}

/**
 * Wrapper function to send a single grade report.
 */
function sendGradesSelect(){
  sendGrades(false);
}

/**
 * Get first non empty row in given column.
 */
function getFirstWrittenRow(col) {
  var sheet = SpreadsheetApp.getActiveSheet(),
      values = sheet.getRange(col + ':'  +col)
                    .getValues();

  for(i = 1; values[i][0] == ''; ++i );

  return ++i;
}

function columnToLetter(col) {
  var i, letter = '';

  for (i = (col - 1) % 26; col > 0; col = (col - i - 1) / 26)  {
    letter = String.fromCharCode(i + 65) + letter;
  }
  return letter;
}

/**
 * Grade reporting function.
 * @param all <boolean> True: To send all reports
 *                      False: Send one report.
 */
function sendGrades(all){
  var i,
      stats = 4,
      sheet = SpreadsheetApp.getActiveSheet(),
      start = getFirstWrittenRow('A'),
      obj = {
        cRanges: {
          header: start,
          lName: 'A',
          fName: 'B',
          email: 'D',
          grade: 'F',
          comment: 'G',
          subGrade: ['H', columnToLetter(sheet.getLastColumn())],
          student: [start + 1, sheet.getLastRow() - stats ]
        },
        appName: appName,
        course : 'COMP 4610',
        assignment: sheet.getName(),
        report: makeGradeTable,
        replyTo: 'wzhou@cs.uml.edu ',
        debugEmail: 'jose.flores.152@gmail.com',
        debug: userProp.getProperty(g_debug_key)
      };

  if(all) {
    // Cancelled multiple student.
    if(showPromptAllStudent(obj) == undefined){
      return undefined;
    }
    // Email all students
    for (i = obj.cRanges.student[0]; i <= obj.cRanges.student[1]; ++i){
      obj.student = i;
      sendEmail(obj);
    }
  } else {
    // Cancelled/ Invalid single student.
    if ((obj.student = showPromptSingleStudent(obj)) == undefined) {
      return undefined;
    }
    // Valid single student
    sendEmail(obj);
  }
}

/**
 * Send an html email.
 */
function sendEmail(obj){
  var sheet = SpreadsheetApp.getActiveSheet(),
    to = (obj.debug == 'true' ? obj.debugEmail : cell(sheet, obj.cRanges.email + obj.student));

  obj.subjectHeading = (obj.course + ' - ' + obj.assignment);

  GmailApp.sendEmail(to, obj.subjectHeading, null, {
    htmlBody: obj.report(obj),
    replyTo: obj.replyTo
  });
}

/**
 * The report generating function.
 * @param    <object>     The values need to generate the report.
 */
function makeGradeTable(obj){

  var tableHead, tableBody,
      sheet = SpreadsheetApp.getActiveSheet(),
      info = {
        'lName': obj.cRanges.lName,
        'fName': obj.cRanges.fName,
        'email': obj.cRanges.email,
        'grade': obj.cRanges.grade,
        'comment': obj.cRanges.comment
      };

  Object.keys(info).forEach(function(key){
    info[key] = cell(sheet, info[key] + obj.student)
  });

  var cssTable = ' style="border-collapse: collapse; border: 1px solid black;"',
      cssCell = ' style="border: 1px solid black; text-align: center;"',
      label  = ['<h1>', obj.subjectHeading, '</h1>',
                '<h3>Results</h3>',
                '<table>',
                    '<tr><th>Name</th><td>', info.lName, ', ', info.fName, '</td></tr>',
                    '<tr><th>Email</th><td>', info.email, '</td></tr>',
                    '<tr><th>Grade</th><td>', info.grade, '%', '</td></tr>',
                '</table>'].join('');

  var subGradeRangeHeader = [obj.cRanges.subGrade[0] + obj.cRanges.header, ':', obj.cRanges.subGrade[1] + obj.cRanges.header].join('');
  var subGradeRangeBody = [obj.cRanges.subGrade[0] + obj.student, ':', obj.cRanges.subGrade[1] + obj.student].join('');

  var subGradesH = sheet.getRange(subGradeRangeHeader),
      subGradesB = sheet.getRange(subGradeRangeBody)

  for (tableHead = tableBody = '', i = 1; i < subGradesH.getNumColumns(); ++i){
    tableHead += ['<th', cssCell, '>', subGradesH.getCell(1,i).getValue(), '</th>'].join('');
    tableBody += ['<td', cssCell, '>', subGradesB.getCell(1,i).getValue(), '</td>'].join('');
  }

  var table = ['<h3>Breakdown</h3>',
               '<table', cssTable, '>',
                   '<thead>',
                       '<tr>', tableHead, '</tr>',
                   '</thead>',
                   '<tbody>',
                       '<tr>', tableBody, '</tr>',
                   '</tbody>',
               '</table>'].join('');

  var comment = ['<h3>Comments</h3>',
             '<p>', escapeHtml(info.comment), '</p>'].join('');

  return [label, table, comment].join('');
}

/**
 * Verifies action on single student.
 */
function showPromptSingleStudent(obj) {
  var sheet = SpreadsheetApp.getActiveSheet(),
      ui = SpreadsheetApp.getUi(),
      query = ui.prompt(obj.appName,
                        'Enter student row number',
                        ui.ButtonSet.OK_CANCEL),
      studentRow = query.getResponseText();

  if (query.getSelectedButton() == ui.Button.OK) {
    if (parseInt(studentRow) &&
        studentRow >= obj.cRanges.student[0] &&
        studentRow <= obj.cRanges.student[1]) {

      var lName = cell(sheet, 'A' + studentRow),
          fName = cell(sheet, 'B' + studentRow);

      return success(obj.appName, ['Emailing student ', lName, ', ', fName, '.'].join(''),
                     studentRow);
    }
    return fail(obj.appName, 'Invalid input for student row.');
  }
  return fail(obj.appName, 'Emailing student was cancelled.');
}

/**
 * Verifies action on all student.
 */
function showPromptAllStudent(obj) {
  var ui = SpreadsheetApp.getUi(),
      result = ui.alert(obj.appName,
                        'Are you sure you want to email all students?',
                        ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    return success(obj.appName, 'All students will be emailed.',
                   true);
  }
  return fail(obj.appName, 'Emailing all students was cancelled.');
}

/**
 * Success dialog.
 */
function success(title, text, value){
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, text, ui.ButtonSet.OK);
  return value;
}

/**
 * Failure dialog.
 */
function fail(title, text){
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, text, ui.ButtonSet.OK);
  return undefined;
}

/**
 * http://stackoverflow.com/questions/1787322/htmlspecialchars-equivalent-in-javascript/4835406#4835406
 *
 * Translates html characters for emailing in html
 */
function escapeHtml(text) {
  var dict = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };

  return text.replace(/[&<>"']/g,
    function(m) {
      return dict[m];
    });
}