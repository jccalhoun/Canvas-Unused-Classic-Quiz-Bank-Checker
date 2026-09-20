/**
* @OnlyCurrentDoc
*/

var BATCH_SIZE = 15;
var MAX_RETRY_ROUNDS = 6;

function isRateLimited(res) {
  var code = res.getResponseCode();
  if (code === 429) return true;
  if (code === 403 && res.getContentText().indexOf('Rate Limit Exceeded') !== -1) return true;
  return false;
}

/**
* @function Fetch a list of URLs concurrently, in chunks, returning raw HTTPResponses
*           in the same order as the input. Retries rate-limited requests with
*           exponential backoff before giving up.
*/
function batchFetchRaw(urlList) {
  var settings = getApiSettings();
  var authHeader = { 'Authorization': 'Bearer ' + settings.token };
  var out = new Array(urlList.length);
  var indexed = urlList.map(function(u, i) { return { i: i, url: u }; });

  chunkArray(indexed, BATCH_SIZE).forEach(function(chunk) {
    var toFetch = chunk.slice();
    var round = 0;
    while (toFetch.length) {
      round++;
      if (round > MAX_RETRY_ROUNDS) {
        throw new Error('Canvas API rate limit: gave up after ' + MAX_RETRY_ROUNDS +
          ' retry rounds on ' + toFetch.length + ' request(s). Try lowering BATCH_SIZE further.');
      }
      var requests = toFetch.map(function(c) {
        return { url: c.url, headers: authHeader, muteHttpExceptions: true };
      });
      var responses = UrlFetchApp.fetchAll(requests);
      var retryNext = [];
      for (var k = 0; k < toFetch.length; k++) {
        var res = responses[k];
        if (isRateLimited(res)) {
          retryNext.push(toFetch[k]);
        } else {
          out[toFetch[k].i] = res;
        }
      }
      if (retryNext.length) {
        var waitMs = Math.min(30000, 1000 * Math.pow(2, round));
        Logger.log('Rate limited on ' + retryNext.length + ' request(s), round ' + round +
          ' — waiting ' + waitMs + 'ms before retry.');
        Utilities.sleep(waitMs);
      }
      toFetch = retryNext;
    }
  });
  return out;
}

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

function chunkArray(arr, size) {
  var chunks = [];
  for (var i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}


/**
* @function Fetch every URL to completion (following pagination) in parallel rounds.
*           Returns, per input URL: an array (for normal paginated list endpoints)
*           or a raw parsed object (for endpoints that don't return a bare array,
*           e.g. the wrapped { quiz_groups: [...] } response).
*/
function batchGetAllPages(startUrls) {
  var settings = getApiSettings();
  var pending = startUrls.map(function(u, i) { return { idx: i, url: u }; });
  var results = startUrls.map(function() { return { isArray: null, data: [] }; });

  while (pending.length) {
    var responses = batchFetchRaw(pending.map(function(p) { return p.url; }));
    var nextPending = [];
    for (var i = 0; i < pending.length; i++) {
      var p = pending[i];
      var res = responses[i];
      var code = res.getResponseCode();
      if (code < 200 || code >= 300) {
        throw new Error('HTTP ' + code + ' on ' + p.url + '\n' + res.getContentText().slice(0, 300));
      }
      var content = res.getContentText();
      var json = content ? JSON.parse(content) : null;
      if (Array.isArray(json)) {
        results[p.idx].isArray = true;
        results[p.idx].data = results[p.idx].data.concat(json);
      } else if (json) {
        results[p.idx].isArray = false;
        results[p.idx].data = json;
      }

      var headers = res.getAllHeaders();
      var link = headers.Link || headers.link;
      if (link) {
        var linkMatch = /<([^>]+)>;\s*rel=["']?next["']?/i.exec(link);
        if (linkMatch) {
          var nextUrl = linkMatch[1];
          // Repair Canvas's occasional host-stripped next-page URL (same issue as before)
          var hostPart = /^https?:\/\/[^\/]+/i.exec(nextUrl);
          if (!hostPart) {
            var pathOnly = nextUrl.replace(/^https?:\/*/i, '/');
            nextUrl = 'https://' + settings.host + pathOnly;
          }
          nextPending.push({ idx: p.idx, url: nextUrl });
        }
      }
    }
    pending = nextPending;
  }

  return results.map(function(r) { return r.data; });
}

/**
* @function Main entry point. Builds the bank-usage report and writes it to a sheet.
*/
function runBankAudit() {
  var t0 = new Date().getTime();
  try {
    var courseId = getCourseId();
    var settings = getApiSettings();
    if (!settings) {
      throw new Error('Canvas API settings are not configured.');
    }
    var host = settings.host;

    var banks = canvasAPI('GET /api/v1/question_banks', {
      'context_type': 'Course',
      'context_id': courseId
    }, ['id', 'title', 'assessment_question_count']);

    if (!banks || !banks.length) {
      SpreadsheetApp.getUi().alert('No question banks found for this course.');
      return;
    }

    // --- Every bank's question IDs, fetched in parallel ---
    var bankUrls = banks.map(function(b) {
      return 'https://' + host + '/api/v1/question_banks/' + b.id + '/questions?per_page=100';
    });
    var bankResults = batchGetAllPages(bankUrls);

    var aqToBank = {};
    banks.forEach(function(b, i) {
      var qs = bankResults[i];
      (Array.isArray(qs) ? qs : []).forEach(function(q) { aqToBank[q.id] = b.id; });
    });
    Logger.log('Bank questions fetched at ' + (new Date().getTime() - t0) + 'ms');

    var usage = {};
    function note(bankId, quizTitle, how) {
      if (!bankId) return;
      usage[bankId] = usage[bankId] || {};
      usage[bankId][quizTitle + ' — ' + how] = true;
    }

    var quizzes = canvasAPI('GET /api/v1/courses/:course_id/quizzes', {
      ':course_id': courseId
    }, ['id', 'title']);

    // --- Every quiz's questions AND groups, fetched in parallel ---
    var quizUrls = [];
    var quizMeta = [];
    quizzes.forEach(function(quiz) {
      quizUrls.push('https://' + host + '/api/v1/courses/' + courseId + '/quizzes/' + quiz.id + '/questions?per_page=100');
      quizMeta.push({ quiz: quiz, type: 'questions' });
      quizUrls.push('https://' + host + '/api/v1/courses/' + courseId + '/quizzes/' + quiz.id + '/groups?per_page=100');
      quizMeta.push({ quiz: quiz, type: 'groups' });
    });
    var quizResults = batchGetAllPages(quizUrls);
    Logger.log('Quiz questions/groups fetched at ' + (new Date().getTime() - t0) + 'ms');

    var byQuiz = {};
    quizMeta.forEach(function(m, i) {
      byQuiz[m.quiz.id] = byQuiz[m.quiz.id] || { title: m.quiz.title };
      byQuiz[m.quiz.id][m.type] = quizResults[i];
    });

    // Quizzes whose /groups response didn't come back as the expected
    // { quiz_groups: [...] } wrapper (or a bare array) need the older
    // per-group fallback — resolved as one more small batched round.
    var fallbackNeeded = [];
    Object.keys(byQuiz).forEach(function(quizId) {
      var entry = byQuiz[quizId];
      var groupsRaw = entry.groups;
      var groupList = (groupsRaw && Array.isArray(groupsRaw.quiz_groups)) ? groupsRaw.quiz_groups
                     : (Array.isArray(groupsRaw) ? groupsRaw : null);
      if (groupList) {
        entry.groupList = groupList;
      } else {
        entry.groupList = [];
        var questions = Array.isArray(entry.questions) ? entry.questions : [];
        var groupIds = [];
        questions.forEach(function(q) {
          if (q.quiz_group_id && groupIds.indexOf(q.quiz_group_id) === -1) {
            groupIds.push(q.quiz_group_id);
          }
        });
        groupIds.forEach(function(gid) {
          fallbackNeeded.push({ quizId: quizId, groupId: gid });
        });
      }
    });

    if (fallbackNeeded.length) {
      var fbUrls = fallbackNeeded.map(function(f) {
        return 'https://' + host + '/api/v1/courses/' + courseId + '/quizzes/' + f.quizId + '/groups/' + f.groupId;
      });
      var fbResults = batchGetAllPages(fbUrls);
      fallbackNeeded.forEach(function(f, i) {
        var g = fbResults[i];
        if (g && !Array.isArray(g)) {
          byQuiz[f.quizId].groupList.push(g);
        }
      });
    }

    Object.keys(byQuiz).forEach(function(quizId) {
      var entry = byQuiz[quizId];
      entry.groupList.forEach(function(g) {
        note(g.assessment_question_bank_id, entry.title, 'linked group "' + (g.name || 'unnamed') + '"');
      });
      (Array.isArray(entry.questions) ? entry.questions : []).forEach(function(q) {
        note(aqToBank[q.assessment_question_id], entry.title, 'question copied in');
      });
    });

    writeBankAuditSheet(banks, usage);
    Logger.log('Total runtime: ' + (new Date().getTime() - t0) + 'ms');
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