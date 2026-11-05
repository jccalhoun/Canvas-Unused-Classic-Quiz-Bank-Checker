/**
* @OnlyCurrentDoc
*
* Flags New Quizzes that pull from archived (deleted) or empty item banks.
* Canvas shows these in the quiz editor as "Item bank no longer exists!" and
* "Invalid Items. Students will not be able to take this assessment...".
*
* Course exports (.imscc) still contain the archived bank's questions, so this
* can only be detected through the live New Quizzes API.
*
* Depends on: getApiSettings(), canvasAPI() (canvasAPI.gs),
*             getCourseId(), batchGetAllPages(), showError() (bankAudit.gs, utils.gs)
*/

function runNewQuizBankCheck() {
  try {
    var courseId = getCourseId();
    var settings = getApiSettings();
    if (!settings) {
      throw new Error('Canvas API settings are not configured.');
    }
    var host = settings.host;

    var quizzes = canvasAPI('GET /api/quiz/v1/courses/:course_id/quizzes', {
      ':course_id': courseId
    }, ['id', 'title']);

    if (!quizzes || !quizzes.length) {
      SpreadsheetApp.getUi().alert('No New Quizzes found for this course.');
      return;
    }

    // --- Every New Quiz's items, fetched in parallel ---
    var itemUrls = quizzes.map(function(q) {
      return 'https://' + host + '/api/quiz/v1/courses/' + courseId + '/quizzes/' + q.id + '/items?per_page=100';
    });
    var itemResults = batchGetAllPages(itemUrls);

    var rows = [];
    var bankSlots = 0;
    quizzes.forEach(function(quiz, i) {
      var items = Array.isArray(itemResults[i]) ? itemResults[i] : [];
      items.forEach(function(item) {
        var e = item.entry;
        var problem = '';
        var bankTitle = '';
        var bankId = '';

        if (item.entry_type === 'Bank') {
          bankSlots++;
          bankTitle = e ? e.title : '';
          bankId = e ? e.id : '';
          if (!e) {
            problem = 'Bank entry missing';
          } else if (e.archived) {
            problem = 'Bank archived (deleted)';
          } else if (!e.item_entry_count) {
            problem = 'Bank is empty';
          }
        } else if (item.entry_type === 'BankEntry') {
          bankId = e ? e.bank_id : '';
          if (!e) {
            problem = 'Bank question missing';
          } else if (e.archived) {
            problem = 'Bank question archived';
          }
        }

        if (problem) {
          rows.push([quiz.title, quiz.id, item.position, bankTitle, bankId, problem]);
        }
      });
    });

    writeNewQuizBankSheet(rows);
    SpreadsheetApp.getUi().alert(
      'New Quiz Bank Check',
      'Checked ' + quizzes.length + ' New Quizzes (' + bankSlots + ' bank slots).\n' +
      (rows.length ? rows.length + ' problem(s) found — see the "New Quiz Bank Check" sheet.' : 'No problems found.'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    showError('New Quiz Bank Check Error', e);
  }
}

function writeNewQuizBankSheet(problemRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'New Quiz Bank Check';
  var sheet = ss.getSheetByName(name);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(name, 0);

  var rows = [['Quiz', 'Assignment ID', 'Item Position', 'Bank Title', 'Bank ID', 'Problem']];
  if (problemRows.length) {
    problemRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
    rows = rows.concat(problemRows);
  } else {
    rows.push(['No problems found', '', '', '', '', '']);
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
}
