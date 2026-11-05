/**
* @OnlyCurrentDoc
*
* New Quizzes item bank audit.
*
* Canvas has no API that lists a course's item banks, so this works from the
* other direction: it reads every New Quiz's items and reports each item bank
* those quizzes point to — which quizzes use it, how, and whether it still works.
* Item banks that no quiz uses cannot be found this way.
*
* Flags banks that are archived (deleted in the Item Banks page). Canvas shows
* those in the quiz editor as "Item bank no longer exists!" / "Invalid Items".
* A course export (.imscc) still contains an archived bank's questions, so this
* can only be detected through the live API.
*
* Depends on: getApiSettings(), canvasAPI() (canvasAPI.gs),
*             getCourseId(), batchGetAllPages() (bankAudit.gs), showError() (utils.gs)
*/

/**
* @function Fetch every New Quiz's items and build a map of the item banks they use.
* @param {String} courseId
* @param {String} host
* @returns {Object} { quizCount, banks: { bankId: bankRecord }, slotProblems: [rows] }
*/
function fetchNewQuizBankUsage(courseId, host) {
  var quizzes = canvasAPI('GET /api/quiz/v1/courses/:course_id/quizzes', {
    ':course_id': courseId
  }, ['id', 'title']) || [];

  var banks = {};
  var slotProblems = [];
  if (!quizzes.length) {
    return { quizCount: 0, banks: banks, slotProblems: slotProblems };
  }

  var itemUrls = quizzes.map(function(q) {
    return 'https://' + host + '/api/quiz/v1/courses/' + courseId + '/quizzes/' + q.id + '/items?per_page=100';
  });
  var itemResults = batchGetAllPages(itemUrls);

  function bankRecord(id) {
    if (!banks[id]) {
      banks[id] = { id: id, title: '', itemCount: '', archived: false, archivedQuestions: 0, uses: {} };
    }
    return banks[id];
  }

  quizzes.forEach(function(quiz, i) {
    var items = Array.isArray(itemResults[i]) ? itemResults[i] : [];
    var individual = {}; // bankId -> number of single questions pulled from it

    items.forEach(function(item) {
      var e = item.entry;

      if (item.entry_type === 'Bank') {
        // A whole bank (or a random sample of it) pulled into the quiz
        if (!e || !e.id) {
          slotProblems.push([quiz.title, quiz.id, item.position, 'Bank slot has no bank data']);
          return;
        }
        var b = bankRecord(String(e.id));
        b.title = e.title || b.title;
        b.itemCount = e.item_entry_count;
        b.archived = b.archived || !!e.archived;
        var sample = item.properties && item.properties.sample_num;
        var how = sample ? 'random ' + sample + ' of ' + e.item_entry_count
                         : 'all ' + e.item_entry_count + ' questions';
        b.uses[quiz.title + ' — ' + how] = true;

      } else if (item.entry_type === 'BankEntry') {
        // A single question added from a bank
        if (!e || !e.bank_id) {
          slotProblems.push([quiz.title, quiz.id, item.position, 'Bank question has no bank data']);
          return;
        }
        var id = String(e.bank_id);
        individual[id] = (individual[id] || 0) + 1;
        if (e.archived) {
          bankRecord(id).archivedQuestions++;
        }
      }
    });

    Object.keys(individual).forEach(function(id) {
      var n = individual[id];
      bankRecord(id).uses[quiz.title + ' — ' + n + ' individual question' + (n === 1 ? '' : 's')] = true;
    });
  });

  return { quizCount: quizzes.length, banks: banks, slotProblems: slotProblems };
}

/**
* @function Status text for one item bank record.
*/
function newQuizBankStatus(b) {
  var s = [];
  if (b.archived) {
    s.push('ARCHIVED (deleted)');
  } else if (b.itemCount === 0) {
    s.push('EMPTY');
  }
  if (b.archivedQuestions) {
    s.push(b.archivedQuestions + ' archived question' + (b.archivedQuestions === 1 ? '' : 's') + ' in use');
  }
  return s.length ? s.join('; ') : 'OK';
}

/**
* @function Main entry point. Builds the New Quizzes item bank report.
*/
function runNewQuizBankAudit() {
  try {
    var courseId = getCourseId();
    var settings = getApiSettings();
    if (!settings) {
      throw new Error('Canvas API settings are not configured.');
    }

    var result = fetchNewQuizBankUsage(courseId, settings.host);
    if (!result.quizCount) {
      SpreadsheetApp.getUi().alert('No New Quizzes found for this course.');
      return;
    }

    var problemCount = writeNewQuizBankSheet(result);
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'New Quiz Bank Audit',
      'Checked ' + result.quizCount + ' New Quizzes using ' + Object.keys(result.banks).length + ' item bank(s).\n' +
      (problemCount ? problemCount + ' problem(s) found — see the "New Quiz Bank Audit" sheet.' : 'No problems found.'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    showError('New Quiz Bank Audit Error', e);
  }
}

/**
* @function Write the New Quizzes report. Problem rows sort to the top.
* @returns {Number} number of problem rows
*/
function writeNewQuizBankSheet(result) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'New Quiz Bank Audit';
  var sheet = ss.getSheetByName(name);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(name, 0);

  var bankList = Object.keys(result.banks).map(function(id) { return result.banks[id]; });
  var problemCount = 0;

  var bankRows = bankList.map(function(b) {
    var status = newQuizBankStatus(b);
    if (status !== 'OK') problemCount++;
    return [
      b.id,
      b.title || '(title not returned — only used for individual questions)',
      b.itemCount,
      status,
      Object.keys(b.uses).join('\n')
    ];
  }).sort(function(a, b) {
    var ap = a[3] === 'OK' ? 1 : 0;
    var bp = b[3] === 'OK' ? 1 : 0;
    return ap - bp || String(a[1]).localeCompare(String(b[1]));
  });

  var slotRows = result.slotProblems.map(function(p) {
    problemCount++;
    return ['(none)', '(unknown bank)', '', p[3], p[0] + ' — item ' + p[2]];
  });

  var rows = [['Bank ID', 'Bank Title', '# Items', 'Status', 'Used By']].concat(slotRows, bankRows);
  if (rows.length === 1) {
    rows.push(['', 'No New Quizzes use an item bank', '', '', '']);
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
  return problemCount;
}
