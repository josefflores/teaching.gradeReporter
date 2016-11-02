/**
 * Jose Flores
 * jose.flores.152@gmail.com
 * 10/20/16
 */

// Defaults
var g_debug_val = 'false';
var appName = 'Grade reporter';
var courseName = 'COMP 4610';
var professorEmail = 'wzhou@cs.uml.edu';

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
 * @param   debug   <String>    The string representation of true or false,
 *                              'true' display debug menu
 *                              'false' display normal menu
 */
function loadMenu(debug) {
    var menu;

    menu = SpreadsheetApp.getUi()
        .createMenu('Teaching');

    // Generate program options
    menu.addItem('Send grades to all student rows', 'sendGradesAll')
        .addItem('Send grade to individual student by row', 'sendGradesSelect')
        .addSeparator();

    // Debug is on - Generate Debug options
    if (debug == 'true') {
        menu.addItem('Turn debug off', 'turnOffDebug')
            .addItem(['Change debug email <',
                    userProp.getProperty(g_debugEmail_key),
                    '>'
                ].join(''),
                'changeDebugEmail')
            .addItem('Reset debug defaults', 'resetDebug');
        //  Debug is off - No debug options
    } else {
        menu.addItem('Turn debug on', 'turnOnDebug');
    }

    // Update UI
    menu.addToUi();
}

/**
 * Sets sheet properties to default values.
 */
function setDefaults() {
    // Default debug email from current user
    if (userProp.getProperty(g_debugEmail_key) === null)
        userProp.setProperty(g_debugEmail_key,
            Session.getActiveUser()
            .getEmail());

    // Default debug state
    if (userProp.getProperty(g_debug_key) === null)
        userProp.setProperty(g_debug_key, g_debug_val);
}

/**
 * Menu Action - Reset sheet properties, and reloads menu with changes.
 */
function resetDebug() {
    //  Delete stored properties
    userProp.deleteProperty(g_debugEmail_key);
    userProp.deleteProperty(g_debug_key);

    //  New default properties, will cause reset
    setDefaults();
    loadMenu(userProp.getProperty(g_debug_key));
}

/**
 * Menu Action - Turns on debug mode, and reloads menu with changes.
 */
function turnOnDebug() {
    userProp.setProperty(g_debug_key, 'true');
    loadMenu(userProp.getProperty(g_debug_key));
}
/**
 * Menu Action - Turns off debug mode, and reloads menu with changes.
 */
function turnOffDebug() {
    userProp.setProperty(g_debug_key, 'false');
    loadMenu(userProp.getProperty(g_debug_key));
}

/**
 * Validates an email for structure.
 *
 * @param   email   <String>    The email to validate
 * @return          <Bool>      true - valid email
 *                              false - invalid email
 */
function validateEmail(email) {
    // Valid email regex from http://emailregex.com/
    var re = /^[-a-z0-9~!$%^&*_=+}{\'?]+(\.[-a-z0-9~!$%^&*_=+}{\'?]+)*@([a-z0-9_][-a-z0-9_]*(\.[-a-z0-9_]+)*\.(aero|arpa|biz|com|coop|edu|gov|info|int|mil|museum|name|net|org|pro|travel|mobi|[a-z][a-z])|([0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}))(:[0-9]{1,5})?$/i;
    return re.test(email);
}

/**
 * Menu Action - Kicks off the debug email change, prompts user for input,
 * validates, and reloads menu with changes.
 */
function changeDebugEmail() {
    var query, email, ui;

    //  Get user interface
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
            userProp.setProperty(g_debugEmail_key,
                email);
            loadMenu(userProp.getProperty(g_debug_key));
            return success(appName,
                'Debug email was changed.',
                userProp.getProperty(g_debugEmail_key));
        }
        // Email is invalid
        return fail(appName,
            'Invalid email.');
    }
    // Canceled Action
    return fail(appName,
        'Debug email change was cancelled.');
}

/**
 * Shorthand for getting value of cell.
 *
 * @param   sheet   <Sheet>     The sheet to extract from.
 * @param   range   <String>    The range to extract.
 * @return                      The value of the cell.
 */
function cell(sheet, range) {
    return sheet.getRange(range)
        .getCell(1, 1)
        .getValue();
}

/**
 * Menu Action - Wrapper function to send all grade reports.
 */
function sendGradesAll() {
    sendGrades(true);
}

/**
 * Menu Action - Wrapper function to send a single grade report.
 */
function sendGradesSelect() {
    sendGrades(false);
}

/**
 * Get first non empty row in a given column.
 *
 * @param   col     <String>    Column identifier.
 * @return                      The first row to have text.
 */
function getFirstWrittenRow(col) {
    var i, sheet, values;

    //  Get column values
    sheet = SpreadsheetApp.getActiveSheet();
    values = sheet.getRange([col,
            ':',
            col
        ].join(''))
        .getValues();

    //  Count rows that are not empty
    for (i = 1; values[i][0] == ''; ++i);

    //  Return next row, the first with text
    return ++i;
}

/**
 * Convert a column number to the letter value.
 *
 * @param   col     <Integer>   The column index.
 * @returns                     The letter equivilance of the column index.
 */
function columnToLetter(col) {
    var i, charOffset, letter;

    letter = '';
    charOffset = 65;
    lettersInAlphabet = 26;

    // Grabbed from stack overflow
    for (i = (col - 1) % lettersInAlphabet; col > 0; col = (col - i - 1) / lettersInAlphabet) {
        letter = String.fromCharCode(i + charOffset) + letter;
    }

    return letter;
}

/**
 * Grade reporting function.
 *
 * @param   all     <boolean>   true: To send all reports
 *                              false: Send one report.
 */
function sendGrades(all) {
    var i, stats, sheet, start, obj;

    /**
     * The number of rows past the last student that text is displayed,
     * We expect to have MEAN MODE MIN MAX
     */
    stats = 4;
    /**
     * Get the first row with text in column A, ie the Header row.
     */
    start = getFirstWrittenRow('A');

    //  The current sheet
    sheet = SpreadsheetApp.getActiveSheet();

    //  Settings to pass on
    obj = {
        //  Positional information columns rows or a range of them as [start, stop]
        cRanges: {
            header: start, // row
            lName: 'A',
            fName: 'B',
            email: 'D',
            grade: 'F',
            comment: 'G',
            //  Range of columns
            subGrade: ['H',
                columnToLetter(sheet.getLastColumn())
            ],
            //  Range of rows
            student: [start + 1,
                sheet.getLastRow() - stats
            ]
        },
        //  Display and emailing properties
        appName: appName,
        course: courseName,
        assignment: sheet.getName(),
        report: makeGradeTable,
        replyTo: professorEmail,
        debugEmail: userProp.getProperty(g_debugEmail_key),
        debug: userProp.getProperty(g_debug_key)
    };

    // If all students should be emailed
    if (all) {
        // Cancelled multiple student.
        if (showPromptAllStudent(obj) == undefined) {
            return undefined;
        }
        // Email all students
        for (i = obj.cRanges.student[0]; i <= obj.cRanges.student[1]; ++i) {
            obj.student = i;
            sendEmail(obj);
        }
    } else {
        // Cancelled/ Invalid single student.
        obj.student = showPromptSingleStudent(obj);
        if (obj.student == undefined) {
            return undefined;
        }
        // Valid single student
        sendEmail(obj);
    }
}

/**
 * Send an html email.
 *
 * @param   obj     <Object>    The settings object.
 */
function sendEmail(obj) {
    var to, sheet;

    //  Get spreadsheet
    sheet = SpreadsheetApp.getActiveSheet();

    //  Determine recipient
    if (obj.debug == 'true') {
        //  Debug mode
        to = obj.debugEmail;
    } else {
        //  Production
        to = cell(sheet, [obj.cRanges.email,
            obj.student
        ].join(''));
    }

    //  Make email heading
    obj.subjectHeading = [
        obj.course,
        ' - ',
        obj.assignment
    ].join('');

    //  Send HTML email
    GmailApp.sendEmail(to,
        obj.subjectHeading,
        null, {
            htmlBody: obj.report(obj),
            replyTo: obj.replyTo
        });
}

/**
 * The report generating function.
 * @param    <object>     The values need to generate the report.
 */
function makeGradeTable(obj) {
    var sheet, info, css, subGradesH, subGradesB, table;

    //  Initialize table object
    table = {
        th: '',
        td: ''
    };

    //  Collect base information for insertion into html
    info = {
        'lName': obj.cRanges.lName,
        'fName': obj.cRanges.fName,
        'email': obj.cRanges.email,
        'grade': obj.cRanges.grade,
        'comment': obj.cRanges.comment
    };

    //  Get spreadsheet
    sheet = SpreadsheetApp.getActiveSheet();

    //  Update info collection for the specified student
    Object.keys(info).forEach(function (key) {
        info[key] = cell(sheet, [
            info[key],
            obj.student
        ].join(''))
    });

    //  Fetch ranges of label heading and scores body
    subGradesH = sheet.getRange([
        obj.cRanges.subGrade[0],
        obj.cRanges.header,
        ':',
        obj.cRanges.subGrade[1],
        obj.cRanges.header
    ].join(''));

    subGradesB = sheet.getRange([
        obj.cRanges.subGrade[0],
        obj.student,
        ':',
        obj.cRanges.subGrade[1],
        obj.student
    ].join(''));

    //  Prepare HTML

    //  Generate score table th and td cells
    css = ' style="border: 1px solid black; text-align: center;"';
    for (i = 1; i < subGradesH.getNumColumns(); ++i) {
        table.th += ['<th', css, '>',
            subGradesH.getCell(1, i)
            .getValue(),
            //' (',
            //cell(sheet, [obj.pointsRow, i].join(''),
            //')',
            '</th>'
        ].join('');

        table.td += ['<td', css, '>',
            subGradesB.getCell(1, i)
            .getValue(),
            '</td>'
        ].join('');
    }

    //  Return formatted email
    return ['<h1>', obj.subjectHeading, '</h1>',
        '<h3>Results</h3>',
        '<table>',
        '<tr><th>Name</th><td>', info.lName, ', ', info.fName, '</td></tr>',
        '<tr><th>Email</th><td>', info.email, '</td></tr>',
        //  Only show grades to two decimals
        '<tr><th>Grade</th><td>', info.grade.toFixed(2), '%', '</td></tr>',
        '</table>',
        '<h3>Breakdown</h3>',
        '<table style="border-collapse: collapse; border: 1px solid black;">',
        '<thead>',
        '<tr>', table.th, '</tr>',
        '</thead>',
        '<tbody>',
        '<tr>', table.td, '</tr>',
        '</tbody>',
        '</table>',
        '<h3>Comments</h3>',
        //  Allow for html to be seen in comments
        '<p>', escapeHtml(info.comment), '</p>'
    ].join('');
}

/**
 * Verifies action on single student.
 *
 * @param   obj     <Object>
 */
function showPromptSingleStudent(obj) {
    var sheet, ui, query, studentRow, lName, fName;

    //  Get spreadsheet UI
    ui = SpreadsheetApp.getUi();

    //  Prompt user for student row number
    query = ui.prompt(obj.appName,
        'Enter student row number',
        ui.ButtonSet.OK_CANCEL);
    studentRow = query.getResponseText();

    //  OK Response
    if (query.getSelectedButton() == ui.Button.OK) {
        //  Valid student
        if (parseInt(studentRow) &&
            studentRow >= obj.cRanges.student[0] &&
            studentRow <= obj.cRanges.student[1]) {

            //  Get spreadsheet
            sheet = SpreadsheetApp.getActiveSheet();

            //  Get Student Name
            lName = cell(sheet, ['A', studentRow].join(''));
            fName = cell(sheet, ['B', studentRow].join(''));

            //  Success Message
            return success(obj.appName, [
                    'Emailing student ',
                    lName,
                    ', ',
                    fName,
                    '.'
                ].join(''),
                studentRow);
        }
        //  Invalid Student
        return fail(obj.appName,
            'Invalid input for student row.');
    }
    //  Cancelled Action
    return fail(obj.appName,
        'Emailing student was cancelled.');
}

/**
 * Verifies action on all student.
 *
 * @param   obj     <Object>    Application information object.
 * @return                      The data that was retrieved.
 */
function showPromptAllStudent(obj) {
    var query, ui;

    //  Get spreadsheet UI
    ui = SpreadsheetApp.getUi();

    // Prompt and retrieve confirmation
    query = ui.alert(obj.appName,
        'Are you sure you want to email all students?',
        ui.ButtonSet.YES_NO);

    //  Send email to all students
    if (query == ui.Button.YES) {
        return success(obj.appName,
            'All students will be emailed.',
            true);
    }

    //  Cancel Action
    return fail(obj.appName,
        'Emailing all students was cancelled.');
}

/**
 * Success dialog.
 *
 * @param   title   <String>    The title of the alert.
 * @param   text    <String>    The body of the alert.
 * @param   value               Any data recieved.
 * @return                      Any data that was passed to value
 */
function success(title, text, value) {
    var ui;

    //  Get spreadsheet UI
    ui = SpreadsheetApp.getUi();

    //  Make success alert
    ui.alert(title,
        text,
        ui.ButtonSet.OK);

    //  Pass value through
    return value;
}

/**
 * Failure dialog.
 *
 * @param   title   <String>    The title of the alert.
 * @param   text    <String>    The body of the alert.
 * @return          undefined   No data passed through.
 */
function fail(title, text) {
    var ui;

    //  Get spreadsheet UI
    ui = SpreadsheetApp.getUi();

    //  Make failure alert
    ui.alert(title,
        text,
        ui.ButtonSet.OK);

    //  Return error
    return undefined;
}

/**
 * http://stackoverflow.com/questions/1787322/htmlspecialchars-equivalent-in-javascript/4835406#4835406
 *
 * Translates html characters for emailing in html
 *
 * @param   text    <String>    An string containing HTML
 * @return          <String>    A string that has had HTML characters encoded.
 */
function escapeHtml(text) {
    var dict = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };

    //  Replace all characters that match keys with their value in dict.
    return text.replace(/[&<>"']/g,
        function (m) {
            return dict[m];
        });
}