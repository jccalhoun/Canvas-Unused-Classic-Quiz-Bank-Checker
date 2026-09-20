/**
* @OnlyCurrentDoc
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Canvas')
    .addItem('Specify Course', 'getCourseDlg')
    .addItem('Run Bank Audit', 'runBankAudit')
    .addSeparator()
    .addItem('Configure API Settings', 'configurationDialog')
    .addItem('Forget API Settings', 'resetApiSettings')
    .addToUi();
}

function getCourseDlg() {
  getCourseDialog();
  return;
}

function getCourseId() {
  var courseId = PropertiesService.getUserProperties().getProperty('courseid');
  if (!courseId) {
    throw new Error('Specify a course first (Canvas menu > Specify Course).');
  }
  return courseId;
}

/**
* @function Main entry point. Builds the bank-usage report and writes it to a sheet.
*/
function runBankAudit() {
  try {
    var courseId = getCourseId();

    var banks = canvasAPI('GET /api/v1/question_banks', {
      'context_type': 'Course',
      'context_id': courseId
    }, ['id', 'title', 'assessment_question_count']);

    if (!banks || !banks.length) {
      SpreadsheetApp.getUi().alert('No question banks found for this course.');
      return;
    }

    // assessment_question_id -> bank id (catches questions copied into quizzes)
    var aqToBank = {};
    banks.forEach(function(b) {
      var qs = canvasAPI('GET /api/v1/question_banks/:id/questions', {
        ':id': b.id
      }, ['id']);
      (qs || []).forEach(function(q) { aqToBank[q.id] = b.id; });
    });

    var usage = {}; // bankId -> { "quiz title | how" : true }
    function note(bankId, quizTitle, how) {
      if (!bankId) return;
      usage[bankId] = usage[bankId] || {};
      usage[bankId][quizTitle + ' — ' + how] = true;
    }

    var quizzes = canvasAPI('GET /api/v1/courses/:course_id/quizzes', {
      ':course_id': courseId
    }, ['id', 'title']);

    (quizzes || []).forEach(function(quiz) {
      var questions = canvasAPI('GET /api/v1/courses/:course_id/quizzes/:quiz_id/questions', {
        ':course_id': courseId,
        ':quiz_id': quiz.id
      }, ['assessment_question_id', 'quiz_group_id']);

      var groups = [];
      try {
        var groupsResp = canvasAPI('GET /api/v1/courses/:course_id/quizzes/:quiz_id/groups', {
          ':course_id': courseId,
          ':quiz_id': quiz.id
        }); // no filter — Canvas wraps this one as { quiz_groups: [...] }, not a bare array
        if (groupsResp && Array.isArray(groupsResp.quiz_groups)) {
          groups = groupsResp.quiz_groups;
        } else if (Array.isArray(groupsResp)) {
          groups = groupsResp; // in case some Canvas versions ever return it bare
        }
        if (!groups.length) {
          throw new Error('No groups returned — falling back to derived group IDs.');
        }
      } catch (e) {
        // Index endpoint not available on this Canvas build — derive from questions instead
        var groupIds = [];
        (questions || []).forEach(function(q) {
          if (q.quiz_group_id && groupIds.indexOf(q.quiz_group_id) === -1) {
            groupIds.push(q.quiz_group_id);
          }
        });
        groups = groupIds.map(function(gid) {
          try {
            return canvasAPI('GET /api/v1/courses/:course_id/quizzes/:quiz_id/groups/:id', {
              ':course_id': courseId,
              ':quiz_id': quiz.id,
              ':id': gid
            });
          } catch (e2) {
            return null;
          }
        }).filter(function(g) { return g; });
      }

      groups.forEach(function(g) {
        note(g.assessment_question_bank_id, quiz.title, 'linked group "' + (g.name || 'unnamed') + '"');
      });
      (questions || []).forEach(function(q) {
        note(aqToBank[q.assessment_question_id], quiz.title, 'question copied in');
      });
    });

    writeBankAuditSheet(banks, usage);
  } catch (e) {
    Logger.log(e);
    showError('Bank Audit Error', e);
  }
}

function writeBankAuditSheet(banks, usage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'Bank Audit';
  var sheet = ss.getSheetByName(name);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(name, 0);

  var rows = [['Bank ID', 'Bank Title', '# Questions', 'Used?', 'Used By']];
  banks.slice().sort(function(a, b) {
    return a.title.localeCompare(b.title) || a.id - b.id;
  }).forEach(function(b) {
    var u = usage[b.id];
    rows.push([
      b.id,
      b.title,
      b.assessment_question_count,
      u ? 'YES' : 'no',
      u ? Object.keys(u).join('\n') : ''
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
}