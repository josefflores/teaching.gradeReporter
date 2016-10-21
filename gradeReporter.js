/**
 * Jose Flores
 * jose.flores.152@gmail.com
 * 10/20/16
 */

/**
 * Adds a Teaching menu to to the spreadsheet menu.
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Teaching')
        .addItem('Send All Grades', 'sendGradesAll')
        .addItem('Send Selected Grade', 'sendGradesSelect')
        .addToUi();
}

/**
 * Shorthand for getting value of cell.
 * @param    sheet    <Sheet>   The sheet to extract from.
 * @param    range    <string>  The range to extract.
 * @return   The value of the cell.
 */
function cell(sheet, range) {
    return sheet.getRange(range)
        .getCell(1, 1)
        .getValue();
}

/**
 * Wrapper function to send all grade reports.
 */
function sendGradesAll() {
    sendGrades(true);
}

/**
 * Wrapper function to send a single grade report.
 */
function sendGradesSelect() {
    sendGrades(false);
}

/**
 * Get first non empty row in given column.
 */
function getFirstWrittenRow(col) {
    var values = SpreadsheetApp.getActiveSheet()
        .getRange(col + ':' + col)
        .getValues();

    for (i = 1; values[i][0] == ''; ++i);

    return ++i;
}

function colMapToLetter(col) {
    var temp, letter = '';
    for (temp = (col - 1) % 26; col > 0; col = (col - temp - 1) / 26) {
        letter = String.fromCharCode(temp + 65) + letter;
    }
    return letter;
}

/**
 * Grade reporting function.
 * @param all <boolean> True: To send all reports
 *                      False: Send one report.
 */
function sendGrades(all) {
    var stats = 4,
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
                subGrade: ['H', colMapToLetter(sheet.getLastColumn())],
                student: [start + 1, sheet.getLastRow() - stats]
            },
            appName: 'Grade reporter',
            course: 'COMP 4610',
            assignment: sheet.getName(),
            report: makeGradeTable,
            replyTo: 'wzhou@cs.uml.edu ',
            debugEmail: 'jose.flores.152@gmail.com',
            debug: false
        };

    if (all) {
        // Cancelled multiple student.
        if (showPromptAllStudent(obj) == undefined) {
            return undefined;
        }
        // Email all students
        for (obj.student = obj.cRanges.student[0]; obj.student <= obj.cRanges.student[1]; ++obj.student) {
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
function sendEmail(obj) {
    var sheet = SpreadsheetApp.getActiveSheet(),
        to = (obj.debug ? obj.debugEmail : cell(sheet, obj.cRanges.email + obj.student));

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
function makeGradeTable(obj) {

    var info,
        subGradesH,
        subGradesB,
        tableHead = '',
        tableBody = '',
        sheet = SpreadsheetApp.getActiveSheet(),
        css = {
            table: ' style="border-collapse: collapse; border: 1px solid black;"',
            cell: ' style="border: 1px solid black; text-align: center;"'
        };

    Object.keys({
        'lName': obj.cRanges.lName,
        'fName': obj.cRanges.fName,
        'email': obj.cRanges.email,
        'grade': obj.cRanges.grade,
        'comment': obj.cRanges.comment
    }).forEach(function (key) {
        info[key] = cell(sheet, info[key] + obj.student)
    });

    subGradesH = sheet.getRange([obj.cRanges.subGrade[0] + obj.cRanges.header, ':',
        obj.cRanges.subGrade[1] + obj.cRanges.header].join(''));

    subGradesB = sheet.getRange([obj.cRanges.subGrade[0] + obj.student, ':',
        obj.cRanges.subGrade[1] + obj.student].join(''));

    for (i = 1; i < subGradesH.getNumColumns(); ++i) {
        tableHead += ['<th', css.cell, '>', subGradesH.getCell(1, i).getValue(), '</th>'].join('');
        tableBody += ['<td', css.cell, '>', subGradesB.getCell(1, i).getValue(), '</td>'].join('');
    }

    return ['<h1>', obj.subjectHeading, '</h1>',
            '<h3>Results</h3>',
            '<table>',
                '<tr><th>Name</th><td>', info.lName, ', ', info.fName, '</td></tr>',
                '<tr><th>Email</th><td>', info.email, '</td></tr>',
                '<tr><th>Grade</th><td>', info.grade, '%', '</td></tr>',
            '</table>',
            '<h3>Breakdown</h3>',
            '<table', css.table, '>',
                '<tr>', tableHead, '</tr>',
                '<tr>', tableBody, '</tr>',
            '</table>',
            '<h3>Comments</h3>',
            '<p>', info.comment, '</p>'].join('');
}

/**
 * Verifies action on single student.
 */
function showPromptSingleStudent(obj) {
    var info,
        sheet = SpreadsheetApp.getActiveSheet(),
        ui = SpreadsheetApp.getUi(),
        query = ui.prompt(obj.appName,
            'Enter student row number',
            ui.ButtonSet.OK_CANCEL),
        studentRow = query.getResponseText();

    if (query.getSelectedButton() == ui.Button.OK) {
        if (parseInt(studentRow) &&
            studentRow >= obj.cRanges.student[0] &&
            studentRow <= obj.cRanges.student[1]) {

            Object.keys({
                'lName': 'A',
                'fName': 'B'
            }).forEach(function (key) {
                info[key] = cell(sheet, info[key] + obj.student)
            });

            return success(obj.appName,
                ['Emailing student ', info['lName'], ', ', info['fName'], '.'].join(''),
                studentRow);
        }
        return fail(obj.appName,
            'Invalid input for student row.');
    }
    return fail(obj.appName,
        'Emailing student was cancelled.');
}

/**
 * Verifies action on all students.
 */
function showPromptAllStudent(obj) {
    var ui = SpreadsheetApp.getUi(),
        result = ui.alert(obj.appName,
            'Are you sure you want to email all students?',
            ui.ButtonSet.YES_NO);

    if (result == ui.Button.YES) {
        return success(obj.appName,
            'All students will be emailed.',
            true);
    }
    return fail(obj.appName,
        'Emailing all students was cancelled.');
}

/**
 * Success dialog.
 */
function success(title, text, value) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(title, text, ui.ButtonSet.OK);
    return value;
}

/**
 * Failure dialog.
 */
function fail(title, text) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(title, text, ui.ButtonSet.OK);
    return undefined;
}